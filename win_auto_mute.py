import json
import os
import sys
from ctypes import (POINTER, WINFUNCTYPE, Structure, byref, c_char, c_int,
                    c_uint, c_ulong, c_void_p, create_string_buffer, pointer,
                    sizeof, windll)
from ctypes.wintypes import MSG

from mute_speakers import mute_all_speakers, mute_current_speaker
from winapi_constants import *


GetDesktopWindow    = windll.user32.GetDesktopWindow
GetWindowRect       = windll.user32.GetWindowRect
GetModuleHandle     = windll.kernel32.GetModuleHandleA
RegisterClassEx     = windll.user32.RegisterClassExA
LoadCursor          = windll.user32.LoadCursorA
CreateWindowEx      = windll.user32.CreateWindowExA
ShowWindow          = windll.user32.ShowWindow
GetMessage          = windll.user32.GetMessageA
TranslateMessage    = windll.user32.TranslateMessage
DispatchMessage     = windll.user32.DispatchMessageA
DefWindowProc       = windll.user32.DefWindowProcA
PostQuitMessage     = windll.user32.PostQuitMessage
LoadImage           = windll.user32.LoadImageA
Shell_NotifyIcon    = windll.Shell32.Shell_NotifyIconA
GetCursorPos        = windll.user32.GetCursorPos
CreatePopupMenu     = windll.user32.CreatePopupMenu
AppendMenu          = windll.user32.AppendMenuA
TrackPopupMenu      = windll.user32.TrackPopupMenu
SetForegroundWindow = windll.user32.SetForegroundWindow
DestroyMenu         = windll.user32.DestroyMenu
SendMessage         = windll.user32.SendMessageA
PostMessage         = windll.user32.PostMessageA
DestroyWindow       = windll.user32.DestroyWindow
IsWindowVisible     = windll.user32.IsWindowVisible

# Task tray icon message
WM_TRAYICON    = (WM_APP + 1)

# Menu ID
ID_TRAY_SETTINGS = 1000
ID_PROCESS_NOW   = 1001
ID_LICENSE       = 1002
ID_TRAY_EXIT     = 1003

# Button ID
ID_MUTE           = 100
ID_SET_VOLUME     = 200
ID_TARGET_SPEAKER = 300
ID_OK             = 400
ID_CANCEL         = 500

WNDPROC = WINFUNCTYPE(c_int, c_void_p, c_int, c_uint, POINTER(c_int))

class WNDCLASSEX(Structure):
    _fields_ = [
        ('cbSize',        c_uint),
        ('style',         c_uint),
        ('lpfnWndProc',   WNDPROC),
        ('cbClsExtra',    c_int),
        ('cbWndExtra',    c_int),
        ('hInstance',     c_void_p),
        ('hIcon',         c_void_p),
        ('hCursor',       c_void_p),
        ('hBrush',        c_void_p),
        ('lpszMenuName',  POINTER(c_char)),
        ('lpszClassName', POINTER(c_char)),
        ('hIconSm',       c_void_p)
    ]

class NOTIFYICONDATA(Structure):
    _fields_ = [
        ('cbSize',           c_uint),
        ('hWnd',             c_void_p),
        ('uID',              c_uint),
        ('uFlags',           c_uint),
        ('uCallbackMessage', c_uint),
        ('hIcon',            c_void_p),
        ('szTip',            c_char * 128),
        ('dwState',          c_uint),
        ('dwStateMask',      c_uint),
        ('szInfo',           c_char * 256),
        ('uTimeout',         c_uint),
        ('szInfoTitle',      c_char * 64),
        ('dwInfoFlags',      c_uint),
        ('guidItem',         c_void_p),
        ('hBalloonIcon',     c_void_p)
    ]

class POINT(Structure):
    _fields_ = [("x", c_ulong),
                ("y", c_ulong)]

class RECT(Structure):
    _fields_ = [("left", c_ulong),
                ("top", c_ulong),
                ("right", c_ulong),
                ("bottom", c_ulong)]

# Task tray icon data
nid = NOTIFYICONDATA()
# Setting values
val_mute    = True  # Mute audio device(s).
val_volume  = False # Set the volume to zero.
val_target  = True  # All speakers are targeted to process.
val_logging = False # Not in the setting window. Need to set manually if you want.
# Window handle values
hmain       = None
hchk_mute   = None
hchk_volume = None
hchk_target = None
hbtn_ok     = None
hbtn_cancel = None


def resource_path() -> str:
    path = ''
    if hasattr(sys, '_MEIPASS'):
        path = os.path.join(sys._MEIPASS, 'resources\\') # type: ignore
    else:
        full_path_name = os.path.abspath(sys.argv[0])
        folder_name = os.path.dirname(full_path_name)
        path = os.path.join(folder_name, 'resources\\')
    return path


def show_open_source_license():
    FILE_LICENCE = 'oss_license.html'
    oss_license_file = os.path.join(resource_path(), FILE_LICENCE)
    if os.path.exists(oss_license_file):
        import webbrowser
        url = 'file:///' + oss_license_file
        webbrowser.open(url)


def is_exist_open_source_license() -> bool:
    FILE_LICENCE = 'oss_license.html'
    oss_license_file = os.path.join(resource_path(), FILE_LICENCE)
    return os.path.exists(oss_license_file)


def load_settings():
    full_path_name = os.path.abspath(sys.argv[0])
    folder_name = os.path.dirname(full_path_name)
    base_file_name = os.path.basename(full_path_name)
    split_file_name = os.path.splitext(base_file_name)

    # json file full path name
    json_file_name = os.path.join(folder_name, split_file_name[0] + '.json')

    # Load settings
    if os.path.exists(json_file_name):
        with open(json_file_name, 'r') as f:
            settings = json.load(f)
    else:
        # Default settings
        settings = {
            'mute': True,
            'volume': False,
            'target': True,
            'logging': False
        }

    return settings


def save_settings():
    full_path_name = os.path.abspath(sys.argv[0])
    folder_name = os.path.dirname(full_path_name)
    base_file_name = os.path.basename(full_path_name)
    split_file_name = os.path.splitext(base_file_name)

    # json file full path name
    json_file_name = os.path.join(folder_name, split_file_name[0] + '.json')

    # Load settings for checking if 'logging' is exist.
    settings = load_settings()
    if 'logging' in settings:
        settings = {
            'mute': val_mute,
            'volume': val_volume,
            'target': val_target,
            'logging': settings['logging']
        }
    else:
        settings = {
            'mute': val_mute,
            'volume': val_volume,
            'target': val_target
        }

    # Save settings
    with open(json_file_name, 'w') as f:
        json.dump(settings, f)


def process():
    global val_mute, val_volume, val_target, val_logging

    if val_target:
        mute_all_speakers(val_mute, val_volume, val_logging)
    else:
        mute_current_speaker(val_mute, val_volume, val_logging)


def WindowProc(hwnd, uMsg, wParam, lParam):
    global nid
    global val_mute, val_volume, val_target
    global hmain, hchk_mute, hchk_volume, hchk_target, hbtn_ok, hbtn_cancel

    if uMsg == WM_CREATE:
        return 0

    elif uMsg == WM_CLOSE:
        if IsWindowVisible(hmain) != 0:
            # Setting window is been operating.
            pass
        else:
            Shell_NotifyIcon(NIM_DELETE, pointer(nid))
            DestroyWindow(hwnd)
        return 0

    elif uMsg == WM_DESTROY:
        PostQuitMessage(0)
        return 0

    elif uMsg == WM_TRAYICON:
        lbytes = bytes(lParam)
        action = int.from_bytes(lbytes, 'little')
        if action == WM_RBUTTONDOWN or action == WM_LBUTTONDOWN:
            # Get current cursor position
            cursor = POINT()
            GetCursorPos(byref(cursor))
            # Create popup menu
            hMenu = CreatePopupMenu()
            AppendMenu(hMenu, MF_STRING, ID_TRAY_SETTINGS, b'Settings')
            AppendMenu(hMenu, MF_STRING, ID_PROCESS_NOW, b'Mute now')
            AppendMenu(hMenu, MF_SEPARATOR, None, None)
            if is_exist_open_source_license():
                AppendMenu(hMenu, MF_STRING, ID_LICENSE, b'Open source licenses')
                AppendMenu(hMenu, MF_SEPARATOR, None, None)
            AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, b'Exit win_auto_mute')
            # Show popup menu
            SetForegroundWindow(hwnd)
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, cursor.x, cursor.y, 0, hwnd, None)
            DestroyMenu(hMenu)
        return 0

    elif uMsg == WM_COMMAND:
        if wParam == ID_TRAY_SETTINGS:
            ShowWindow(hwnd, SW_SHOWNORMAL)
            pass
        elif wParam == ID_PROCESS_NOW:
            process()
            pass
        elif wParam == ID_LICENSE:
            show_open_source_license()
            pass
        elif wParam == ID_TRAY_EXIT:
            PostMessage(hwnd, WM_CLOSE, 0, 0)
            pass
        else:
            b = bytes(lParam)
            h = int.from_bytes(b, 'little')
            if h == hchk_mute:
                pass
            elif h == hchk_volume:
                pass
            elif h == hchk_target:
                pass
            elif h == hbtn_ok:
                state_mute   = SendMessage(hchk_mute, BM_GETCHECK, 0, 0)
                state_volume = SendMessage(hchk_volume, BM_GETCHECK, 0, 0)
                state_target = SendMessage(hchk_target, BM_GETCHECK, 0, 0)
                val_mute   = (state_mute   == BST_CHECKED)
                val_volume = (state_volume == BST_CHECKED)
                val_target = (state_target == BST_CHECKED)
                save_settings()
                ShowWindow(hmain, SW_HIDE)
                pass
            elif h == hbtn_cancel:
                ShowWindow(hmain, SW_HIDE)
                pass
        return 0

    elif uMsg == WM_ENDSESSION:
        process()
        return 0

    return DefWindowProc(hwnd, uMsg, wParam, lParam)


def WinMain(class_name, title_name):
    global nid
    global val_mute, val_volume, val_target, val_logging
    global hmain, hchk_mute, hchk_volume, hchk_target, hbtn_ok, hbtn_cancel

    hInstance = GetModuleHandle(None)

    # Window Class name
    wnd_class_name = create_string_buffer(class_name)

    # Window Class attribute
    wcex = WNDCLASSEX()
    wcex.cbSize = sizeof(WNDCLASSEX)
    wcex.style = CS_HREDRAW | CS_VREDRAW
    wcex.lpfnWndProc = WNDPROC(WindowProc)
    wcex.cbClsExtra = 0
    wcex.cbWndExtra = 0
    wcex.hInstance = hInstance
    wcex.hIcon = 0
    wcex.hCursor = LoadCursor(None, IDC_ARROW)
    wcex.hBrush = COLOR_WINDOW
    wcex.lpszMenuName = None
    wcex.lpszClassName = wnd_class_name
    wcex.hIconSm = 0

    # Register Window Class
    regex = RegisterClassEx(pointer(wcex))

    # Window Title
    win_title = create_string_buffer(title_name)

    # Get desktop size
    desktop = windll.user32.GetDesktopWindow()
    rect = RECT()
    GetWindowRect(desktop, byref(rect))
    desktop_width = rect.right - rect.left
    desktop_height = rect.bottom - rect.top

    # Create Window
    width = 350
    height = 200
    hmain = CreateWindowEx(
        0,
        wnd_class_name,
        win_title,
        (WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX),
        (desktop_width - width) // 2, # x0
        (desktop_height - height) // 2, # y0
        width, # cx
        height, # cy
        None,
        None,
        hInstance,
        None)

    if hmain == None:
        import sys
        sys.exit()

    # Check box
    hchk_mute = CreateWindowEx(
        0,
        b'BUTTON',
        b'Mute audio device(s).',
        WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
        10,
        10,
        300,
        20,
        hmain,
        ID_MUTE,
        hInstance,
        None
    )
    hchk_volume = CreateWindowEx(
        0,
        b'BUTTON',
        b'Set the volume to zero.',
        WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
        10,
        40,
        300,
        20,
        hmain,
        ID_SET_VOLUME,
        hInstance,
        None
    )
    hchk_target = CreateWindowEx(
        0,
        b'BUTTON',
        b'All speakers are targeted to process.',
        WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
        10,
        70,
        300,
        20,
        hmain,
        ID_TARGET_SPEAKER,
        hInstance,
        None
    )

    # OK Button
    hbtn_ok = CreateWindowEx(
        0,
        b'BUTTON',
        b'OK',
        WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        30,
        110,
        100,
        30,
        hmain,
        ID_OK,
        hInstance,
        None
    )
    # Cancel Button
    hbtn_cancel = CreateWindowEx(
        0,
        b'BUTTON',
        b'Cancel',
        WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        200,
        110,
        100,
        30,
        hmain,
        ID_CANCEL,
        hInstance,
        None
    )

    # Task tray icon file
    icon_file = os.path.join(resource_path(), 'mute_t.ico')
    # to byes
    icon_file = icon_file.encode('utf-8')

    # Task tray icon data
    nid.cbSize = sizeof(NOTIFYICONDATA)
    nid.hWnd = hmain
    nid.uID = 1
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP # | NIIF_LARGE_ICON
    nid.uCallbackMessage = WM_TRAYICON
    nid.hIcon = LoadImage(None, icon_file, IMAGE_ICON, 0, 0, LR_LOADFROMFILE)
    nid.szTip = b'win_auto_mute'
    Shell_NotifyIcon(NIM_ADD, pointer(nid))

    # Load settings
    settings = load_settings()
    if 'mute' in settings:
        val_mute = settings['mute']
    if 'volume' in settings:
        val_volume = settings['volume']
    if 'target' in settings:
        val_target = settings['target']
    SendMessage(hchk_mute,   BM_SETCHECK, (BST_CHECKED if val_mute   else BST_UNCHECKED), 0)
    SendMessage(hchk_volume, BM_SETCHECK, (BST_CHECKED if val_volume else BST_UNCHECKED), 0)
    SendMessage(hchk_target, BM_SETCHECK, (BST_CHECKED if val_target else BST_UNCHECKED), 0)
    if 'logging' in settings:
        val_logging = settings['logging']
    pass

    # Window Message Structure
    msg = MSG()

    # Message Loop
    while True:
        ret = GetMessage(pointer(msg), None, 0, 0)
        if ret == 0:
            break
        if ret != -1:
            TranslateMessage(pointer(msg))
            DispatchMessage(pointer(msg))
    pass


if __name__ == '__main__':
    WinMain(b'win_auto_mute', b'win_auto_mute')

