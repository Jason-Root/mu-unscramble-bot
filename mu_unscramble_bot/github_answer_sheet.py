from __future__ import annotations

from dataclasses import dataclass
import base64
import json
import urllib.error
import urllib.parse
import urllib.request


API_BASE_URL = "https://api.github.com"


@dataclass(frozen=True, slots=True)
class GitHubAnswerSheetConfig:
    repository: str
    branch: str
    path: str
    token: str | None = None
    sync_interval_seconds: float = 30.0
    commit_message: str = "Update community answer sheet"


@dataclass(frozen=True, slots=True)
class GitHubFileSnapshot:
    text: str
    sha: str | None


class GitHubAnswerSheetClient:
    def __init__(self, config: GitHubAnswerSheetConfig) -> None:
        self.config = config

    def fetch(self) -> GitHubFileSnapshot:
        url = self._contents_url(with_ref=True)
        request = urllib.request.Request(url, headers=self._headers())
        try:
            with urllib.request.urlopen(request, timeout=10.0) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                return GitHubFileSnapshot(text="", sha=None)
            raise

        content = str(payload.get("content", "") or "")
        encoding = str(payload.get("encoding", "") or "")
        sha = str(payload.get("sha", "") or "") or None
        if encoding == "base64" and content:
            decoded = base64.b64decode(content.encode("utf-8"))
            return GitHubFileSnapshot(text=decoded.decode("utf-8"), sha=sha)
        return GitHubFileSnapshot(text="", sha=sha)

    def push(self, text: str, *, sha: str | None) -> str | None:
        payload: dict[str, object] = {
            "message": self.config.commit_message,
            "content": base64.b64encode(text.encode("utf-8")).decode("ascii"),
            "branch": self.config.branch,
        }
        if sha:
            payload["sha"] = sha

        request = urllib.request.Request(
            self._contents_url(with_ref=False),
            data=json.dumps(payload).encode("utf-8"),
            headers=self._headers(),
            method="PUT",
        )
        with urllib.request.urlopen(request, timeout=15.0) as response:
            result = json.loads(response.read().decode("utf-8"))
        content_info = result.get("content", {}) if isinstance(result, dict) else {}
        if isinstance(content_info, dict):
            return str(content_info.get("sha", "") or "") or None
        return None

    def _contents_url(self, *, with_ref: bool) -> str:
        repo = self.config.repository.strip().strip("/")
        path = urllib.parse.quote(self.config.path.strip().lstrip("/"), safe="/")
        url = f"{API_BASE_URL}/repos/{repo}/contents/{path}"
        if with_ref and self.config.branch.strip():
            branch = urllib.parse.quote(self.config.branch.strip(), safe="")
            url = f"{url}?ref={branch}"
        return url

    def _headers(self) -> dict[str, str]:
        headers = {
            "Accept": "application/vnd.github+json",
            "User-Agent": "mu-unscramble-bot",
            "X-GitHub-Api-Version": "2022-11-28",
        }
        token = (self.config.token or "").strip()
        if token:
            headers["Authorization"] = f"Bearer {token}"
        return headers
