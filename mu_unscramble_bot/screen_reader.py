from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import threading
from typing import Any

import cv2
import mss
import numpy as np
from rapidocr_onnxruntime import RapidOCR

from mu_unscramble_bot.config import BotConfig
from mu_unscramble_bot.window_target import get_target_window


@dataclass(slots=True)
class CaptureResult:
    region: dict[str, int]
    frame: np.ndarray
    mask: np.ndarray
    variants: dict[str, np.ndarray]
    lines: list[str]


class YellowTextReader:
    def __init__(self, config: BotConfig) -> None:
        self.config = config
        self._ocr = RapidOCR()
        self._screen: Any | None = None
        self._screen_owner_thread_id: int | None = None
        self._warm_fast_recognizer()

    def close(self) -> None:
        if self._screen is None:
            return
        try:
            self._screen.close()
        finally:
            self._screen = None
            self._screen_owner_thread_id = None

    def read_from_screen(self) -> CaptureResult:
        screen = self._ensure_screen()
        region = self._resolve_live_region(screen)
        frame = np.array(screen.grab(region), dtype=np.uint8)[:, :, :3]
        return self._analyze_frame(frame, region)

    def read_from_image(self, image_path: str | Path) -> CaptureResult:
        image = cv2.imread(str(image_path))
        if image is None:
            raise FileNotFoundError(f"Could not read image: {image_path}")

        region = self._resolve_region(width=image.shape[1], height=image.shape[0], left=0, top=0)
        crop = image[
            region["top"] : region["top"] + region["height"],
            region["left"] : region["left"] + region["width"],
        ]
        return self._analyze_frame(crop, region)

    def _resolve_region(self, width: int, height: int, left: int, top: int) -> dict[str, int]:
        center_x = left + (width // 2) + self.config.center_offset_x
        center_y = top + (height // 2) + self.config.center_offset_y

        region_left = max(left, int(center_x - (self.config.capture_width / 2)))
        region_top = max(top, int(center_y - (self.config.capture_height / 2)))

        max_right = left + width
        max_bottom = top + height
        region_right = min(max_right, region_left + self.config.capture_width)
        region_bottom = min(max_bottom, region_top + self.config.capture_height)

        if region_right - region_left < self.config.capture_width:
            region_left = max(left, region_right - self.config.capture_width)
        if region_bottom - region_top < self.config.capture_height:
            region_top = max(top, region_bottom - self.config.capture_height)

        return {
            "left": int(region_left),
            "top": int(region_top),
            "width": int(region_right - region_left),
            "height": int(region_bottom - region_top),
        }

    def _resolve_live_region(self, screen: Any) -> dict[str, int]:
        if self.config.capture_source.lower() == "window":
            window = get_target_window(self.config)
            return self._resolve_region(
                width=window.width,
                height=window.height,
                left=window.left,
                top=window.top,
            )

        monitor = screen.monitors[self.config.monitor_index]
        return self._resolve_region(
            width=monitor["width"],
            height=monitor["height"],
            left=monitor["left"],
            top=monitor["top"],
        )

    def _ensure_screen(self) -> Any:
        current_thread_id = threading.get_ident()
        if self._screen is not None and self._screen_owner_thread_id == current_thread_id:
            return self._screen

        self.close()
        self._screen = mss.mss()
        self._screen_owner_thread_id = current_thread_id
        return self._screen

    def _analyze_frame(self, frame: np.ndarray, region: dict[str, int]) -> CaptureResult:
        mask = self._yellow_mask(frame)
        ocr_frame, ocr_mask = self._crop_to_mask_bounds(frame, mask)
        variants = self._build_variants(ocr_frame, ocr_mask)
        lines = self._extract_lines(frame, mask, variants)
        return CaptureResult(region=region, frame=frame, mask=mask, variants=variants, lines=lines)

    def _yellow_mask(self, frame: np.ndarray) -> np.ndarray:
        hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
        mask = cv2.inRange(
            hsv,
            np.array(self.config.yellow_hsv_lower, dtype=np.uint8),
            np.array(self.config.yellow_hsv_upper, dtype=np.uint8),
        )
        mask = cv2.medianBlur(mask, 3)
        kernel = np.ones((2, 2), np.uint8)
        return cv2.dilate(mask, kernel, iterations=self.config.mask_dilate_iterations)

    def _build_variants(self, frame: np.ndarray, mask: np.ndarray) -> dict[str, np.ndarray]:
        yellow_only = cv2.bitwise_and(frame, frame, mask=mask)
        contrast = np.full_like(frame, 255)
        contrast[mask > 0] = (0, 0, 0)

        scale = 2
        return {
            "raw": cv2.resize(frame, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC),
            "yellow_only": cv2.resize(
                yellow_only,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_CUBIC,
            ),
            "contrast": cv2.resize(
                contrast,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_CUBIC,
            ),
        }

    def _extract_lines(
        self,
        frame: np.ndarray,
        mask: np.ndarray,
        variants: dict[str, np.ndarray],
    ) -> list[str]:
        fast_lines = self._extract_lines_from_strips(frame, mask)
        if fast_lines:
            return fast_lines
        return self._extract_lines_with_detector(variants)

    def _extract_lines_from_strips(self, frame: np.ndarray, mask: np.ndarray) -> list[str]:
        strips = self._extract_line_strips(frame, mask)
        if not strips:
            return []

        rec_res, _ = self._ocr.text_rec([strip for _, strip in strips])
        candidate_lines: list[tuple[int, str, float]] = []
        seen_keys: set[str] = set()
        for (top, _), (text, score) in zip(strips, rec_res):
            cleaned = self._clean_text(text)
            if not cleaned:
                continue
            if score < self.config.min_ocr_confidence and len(cleaned) < 10:
                continue

            dedupe_key = re.sub(r"[^a-z0-9]+", "", cleaned.lower())
            if not dedupe_key or dedupe_key in seen_keys:
                continue

            seen_keys.add(dedupe_key)
            candidate_lines.append((top, cleaned, float(score)))

        candidate_lines.sort(key=lambda item: (item[0], -item[2]))
        return [text for _, text, _ in candidate_lines]

    def _extract_lines_with_detector(self, variants: dict[str, np.ndarray]) -> list[str]:
        variant_order = ["yellow_only", "contrast", "raw"]
        candidate_lines: list[tuple[float, int, str, float]] = []
        seen_keys: set[str] = set()

        for priority, variant_name in enumerate(variant_order):
            result, _ = self._ocr(variants[variant_name])
            if not result:
                continue

            variant_added = 0
            for box, text, score in result:
                cleaned = self._clean_text(text)
                if not cleaned:
                    continue
                if score < self.config.min_ocr_confidence and len(cleaned) < 10:
                    continue

                dedupe_key = re.sub(r"[^a-z0-9]+", "", cleaned.lower())
                if not dedupe_key or dedupe_key in seen_keys:
                    continue

                seen_keys.add(dedupe_key)
                avg_y = float(sum(point[1] for point in box) / len(box))
                candidate_lines.append((avg_y, priority, cleaned, float(score)))
                variant_added += 1

            if variant_added and self._looks_like_puzzle_text([text for _, _, text, _ in candidate_lines]):
                break

        candidate_lines.sort(key=lambda item: (item[0], item[1], -item[3]))
        return [text for _, _, text, _ in candidate_lines]

    def _extract_line_strips(self, frame: np.ndarray, mask: np.ndarray) -> list[tuple[int, np.ndarray]]:
        merged = cv2.morphologyEx(
            mask,
            cv2.MORPH_CLOSE,
            cv2.getStructuringElement(cv2.MORPH_RECT, (21, 5)),
            iterations=2,
        )
        contours, _ = cv2.findContours(merged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        strips: list[tuple[int, np.ndarray]] = []
        min_block_top = int(frame.shape[0] * 0.1)
        min_block_width = max(80, int(frame.shape[1] * 0.12))
        max_center_distance = frame.shape[1] * 0.38
        frame_center_x = frame.shape[1] / 2

        for contour in contours:
            x, y, width, height = cv2.boundingRect(contour)
            block_center_x = x + (width / 2)
            if width < min_block_width or height < 10:
                continue
            if y + height < min_block_top:
                continue
            if abs(block_center_x - frame_center_x) > max_center_distance:
                continue

            block = frame[y : y + height, x : x + width]
            block_mask = mask[y : y + height, x : x + width]
            min_pixels = max(10, width // 50)
            for band_top, band_bottom in self._find_row_bands(block_mask, min_pixels=min_pixels):
                pad = 3
                top = max(0, band_top - pad)
                bottom = min(height, band_bottom + pad)
                strip = block[top:bottom, :]
                if strip.shape[0] < 10 or strip.shape[1] < 60:
                    continue
                strips.append((y + top, strip))

        strips.sort(key=lambda item: item[0])
        return strips[:12]

    @staticmethod
    def _crop_to_mask_bounds(frame: np.ndarray, mask: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
        points = cv2.findNonZero(mask)
        if points is None:
            return frame, mask

        x, y, width, height = cv2.boundingRect(points)
        margin = 12
        left = max(0, x - margin)
        top = max(0, y - margin)
        right = min(frame.shape[1], x + width + margin)
        bottom = min(frame.shape[0], y + height + margin)
        return frame[top:bottom, left:right], mask[top:bottom, left:right]

    @staticmethod
    def _looks_like_puzzle_text(lines: list[str]) -> bool:
        if not lines:
            return False
        combined = " ".join(lines).lower()
        return any(
            token in combined
            for token in ("round", "hint", "difficulty", "guessed word", "unscramb")
        )

    @staticmethod
    def _find_row_bands(mask: np.ndarray, *, min_pixels: int) -> list[tuple[int, int]]:
        rows = (mask > 0).sum(axis=1)
        bands: list[tuple[int, int]] = []
        start: int | None = None
        for index, value in enumerate(rows):
            if value > min_pixels and start is None:
                start = index
                continue
            if value <= min_pixels and start is not None:
                if index - start >= 5:
                    bands.append((start, index))
                start = None
        if start is not None and len(rows) - start >= 5:
            bands.append((start, len(rows)))
        return bands

    def _warm_fast_recognizer(self) -> None:
        blank = np.full((24, 160, 3), 255, dtype=np.uint8)
        try:
            self._ocr.text_rec(blank)
        except Exception:
            pass

    @staticmethod
    def _clean_text(text: str) -> str:
        text = text.replace("|", "I")
        text = text.replace("—", "-")
        return re.sub(r"\s+", " ", text).strip()
