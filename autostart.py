import sys
import os
import winreg

REG_KEY  = r'Software\Microsoft\Windows\CurrentVersion\Run'
REG_NAME = '서서니 메모'

_DIR = os.path.dirname(os.path.abspath(__file__))


def _cmd():
    if getattr(sys, 'frozen', False):  # PyInstaller exe
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{os.path.join(_DIR, "main.py")}"'


def is_enabled():
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        winreg.QueryValueEx(k, REG_NAME)
        winreg.CloseKey(k)
        return True
    except OSError:
        return False


def refresh_if_enabled():
    """자동 시작이 등록되어 있으면 현재 실행 경로로 갱신한다."""
    if is_enabled():
        set_enabled(True)


def set_enabled(enable: bool):
    k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_SET_VALUE)
    if enable:
        winreg.SetValueEx(k, REG_NAME, 0, winreg.REG_SZ, _cmd())
    else:
        try:
            winreg.DeleteValue(k, REG_NAME)
        except OSError:
            pass
    winreg.CloseKey(k)
