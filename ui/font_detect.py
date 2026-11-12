"""이 PC에 설치된 글꼴 파일 찾기 (Windows 글꼴 등록 정보 + 글꼴 폴더).

Windows용 Qt는 앱을 켤 때 받은 글꼴 목록을 실행 중에 새로 고치지 않는다. 그래서 앱을 켜 둔 채 글꼴을 설치하면
Qt 목록에는 안 보인다. 여기서 Windows가 기록해 둔 글꼴 파일을 직접 찾아, 앱이 그 파일을 불러 쓰게 한다.
  · 내 계정에만 설치(관리자 권한 없이): HKCU\\...\\Fonts, %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts
  · PC 전체에 설치: HKLM\\...\\Fonts, %WINDIR%\\Fonts
"""
import os
import sys

_FONTS_KEY = r'Software\Microsoft\Windows NT\CurrentVersion\Fonts'
_EXTS = ('.ttf', '.otf', '.ttc')


def _windows_fonts_dir():
    return os.path.join(os.environ.get('WINDIR') or r'C:\Windows', 'Fonts')


def _user_fonts_dir():
    base = os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), 'AppData', 'Local')
    return os.path.join(base, 'Microsoft', 'Windows', 'Fonts')


def _registry_files():
    """Windows 글꼴 등록 정보에 적힌 글꼴 파일 경로들 (Windows가 아니면 빈 목록)."""
    if sys.platform != 'win32':
        return []
    try:
        import winreg
    except ImportError:
        return []
    out = []
    for root, base_dir in ((winreg.HKEY_CURRENT_USER, _user_fonts_dir()),
                           (winreg.HKEY_LOCAL_MACHINE, _windows_fonts_dir())):
        try:
            key = winreg.OpenKey(root, _FONTS_KEY)
        except OSError:
            continue
        try:
            i = 0
            while True:
                try:
                    _name, value, _t = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                if isinstance(value, str) and value.lower().endswith(_EXTS):
                    out.append(value if os.path.isabs(value) else os.path.join(base_dir, value))
        finally:
            winreg.CloseKey(key)
    return out


def _folder_files():
    """글꼴 폴더 두 곳의 파일 (등록 정보가 비어 있어도 찾게)."""
    out = []
    if sys.platform != 'win32':
        return out
    for d in (_user_fonts_dir(), _windows_fonts_dir()):
        try:
            out.extend(os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(_EXTS))
        except OSError:
            pass
    return out


def installed_font_files():
    """이 PC에 설치된 글꼴 파일 경로 (실제로 있는 것만, 중복 없이)."""
    seen, out = set(), []
    for p in _registry_files() + _folder_files():
        k = os.path.normcase(os.path.abspath(p))
        if k in seen or not os.path.isfile(p):
            continue
        seen.add(k)
        out.append(p)
    return out


def find_files(stems, files=None):
    """파일 이름(확장자 뺀 것, 대소문자 무시)으로 찾기 → {이름: 경로}. 같은 이름이면 .ttf를 먼저."""
    files = installed_font_files() if files is None else files
    want = {s.lower(): s for s in stems}
    found = {}
    for p in sorted(files, key=lambda p: (not p.lower().endswith('.ttf'), p)):
        stem = os.path.splitext(os.path.basename(p))[0].lower()
        if stem in want and want[stem] not in found:
            found[want[stem]] = p
    return found
