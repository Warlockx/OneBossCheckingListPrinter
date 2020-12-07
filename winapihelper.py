import time

import numpy as np
import psutil
import win32api
import win32con
import win32gui
from subprocess import PIPE
import pywintypes
import win32process
from string import printable


def get_rdp_pid():
    connections = psutil.net_connections('inet')
    for connection in connections:
        if connection.raddr != () and connection.raddr.port == 10003 and connection.status == "ESTABLISHED":
            return connection.pid
    return 0


def open_rdp():
    return psutil.Popen(['mstsc', 'oneboss.rdp', '/w:1350', '/h:700'], stdout=PIPE)


def get_handle(_pid):
    try:
        parent = win32gui.FindWindow(None, 'oneboss - erp.oneboss.com.br:10003 - Conexão de Área de Trabalho Remota')
        win32gui.SetWindowPos(parent, 0, -1366, 0, 1366, 768, 0x0010)
        win32gui.SetWindowPos(parent, 0, -1366, 0, 1366, 768, 0x0001)
        child_windows = []

        def check_child_process(hwnd, param):
            # name = win32gui.GetWindowText(hwnd)
            child_windows.append(hwnd)
        win32gui.EnumChildWindows(parent, check_child_process, None)

    except pywintypes.error:
        return [0, 0]
    return child_windows[2], child_windows[4]


def kill_inactive_windows():
    while True:
        try:
            window = win32gui.FindWindow(None, 'oneboss - erp.oneboss.com.br:10003 - '
                                               'Conexão de Área de Trabalho Remota')
            if window == 0:
                break
            threads = win32process.GetWindowThreadProcessId(window)
            for thread in threads:
                try:
                    proc = psutil.Process(thread)
                    proc.kill()
                except psutil.NoSuchProcess:
                    pass
        except pywintypes.error:
            break


def kill_rdp_error_windows():
    while True:
        try:
            window = win32gui.FindWindow(None, 'Conexão de Área de Trabalho Remota')
            if window == 0:
                break
            threads = win32process.GetWindowThreadProcessId(window)
            for thread in threads:
                try:
                    proc = psutil.Process(thread)
                    proc.kill()
                except psutil.NoSuchProcess:
                    pass

        except pywintypes.error:
            break


def restore_window_if_minimized(handle):
    window = win32gui.FindWindow(None, 'oneboss - erp.oneboss.com.br:10003 - '
                                       'Conexão de Área de Trabalho Remota')
    is_minimized = win32gui.IsIconic(window)
    if is_minimized:
        win32gui.ShowWindow(window, win32con.SW_SHOWNOACTIVATE)


def make_l_param(window_size, coord):
    h, w = window_size
    print('clicked at {}'.format(coord))
    return (coord[1] << 16) | coord[0]


def get_window_size(handle):
    rect = win32gui.GetWindowRect(handle)
    w = rect[2] - rect[0]
    h = rect[3] - rect[1]
    return h, w


def keyboard_click(handle, key):
    if isinstance(key, str) and key in printable:
        key = win32api.MapVirtualKey(ord(key), 2)
        win32api.PostMessage(handle, win32con.WM_KEYDOWN, key)
        win32api.PostMessage(handle, win32con.WM_CHAR,  key)
        win32api.PostMessage(handle, win32con.WM_KEYUP, key)
    else:
        win32api.PostMessage(handle, win32con.WM_KEYDOWN, key, 0)
        win32api.PostMessage(handle, win32con.WM_KEYUP,  key, 0)


def write_text(handle, text):
    for key in text:
        keyboard_click(handle, key)


def mouse_click(handle, coords, offset=(0, 0)):
    best_coords = max(coords[0])

    w, h = coords[1].shape[::-1]
    best_coords = (int(best_coords[0] + w / 2)+offset[0],
                   int(best_coords[1] + h / 2)+offset[1])

    coords_long = win32api.MAKELONG(best_coords[0], best_coords[1])
    time.sleep(.5)

    win32api.PostMessage(handle, win32con.WM_MOUSEACTIVATE)
    win32api.PostMessage(handle, win32con.WM_LBUTTONDOWN, 1, coords_long)
    win32api.PostMessage(handle, win32con.WM_LBUTTONUP, 0, coords_long)
