from __future__ import annotations

import ctypes
from ctypes import wintypes


PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
TOKEN_QUERY = 0x0008
TOKEN_ELEVATION_CLASS = 20


class TOKEN_ELEVATION(ctypes.Structure):
    _fields_ = [("TokenIsElevated", wintypes.DWORD)]


def is_current_process_elevated() -> bool | None:
    kernel32 = ctypes.windll.kernel32
    current_pid = int(kernel32.GetCurrentProcessId())
    return is_pid_elevated(current_pid)


def is_pid_elevated(pid: int) -> bool | None:
    kernel32 = ctypes.windll.kernel32
    process = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not process:
        return None
    return _is_process_handle_elevated(process, close_handle=True)


def get_window_pid(window: object) -> int | None:
    hwnd = int(getattr(window, "_hWnd", 0) or 0)
    if not hwnd:
        return None

    user32 = ctypes.windll.user32
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value or 0) or None


def _is_process_handle_elevated(process: int, *, close_handle: bool) -> bool | None:
    advapi32 = ctypes.windll.advapi32
    kernel32 = ctypes.windll.kernel32
    token = wintypes.HANDLE()

    try:
        if not advapi32.OpenProcessToken(process, TOKEN_QUERY, ctypes.byref(token)):
            return None

        elevation = TOKEN_ELEVATION()
        return_length = wintypes.DWORD()
        ok = advapi32.GetTokenInformation(
            token,
            TOKEN_ELEVATION_CLASS,
            ctypes.byref(elevation),
            ctypes.sizeof(elevation),
            ctypes.byref(return_length),
        )
        if not ok:
            return None
        return bool(elevation.TokenIsElevated)
    finally:
        if token:
            kernel32.CloseHandle(token)
        if close_handle and process:
            kernel32.CloseHandle(process)
