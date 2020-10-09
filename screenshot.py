import pywintypes
import win32gui
import win32ui
import win32con
import numpy as np


class Screenshot:
    def __init__(self, handle):
        self.handle = handle
        self.rect = None
        self.current_img = None
        self.current_npimg = None

    def get_rect(self, rect):
        self.rect = [(rect[2] - rect[0]), (rect[3] - rect[1])]

    def get_image(self, save):
        try:
            self.get_rect(win32gui.GetWindowRect(self.handle))
            window_dc = win32gui.GetWindowDC(self.handle)
            dc_obj = win32ui.CreateDCFromHandle(window_dc)
            compatible_dc = dc_obj.CreateCompatibleDC()
            data_bitmap = win32ui.CreateBitmap()
            data_bitmap.CreateCompatibleBitmap(dc_obj, self.rect[0], self.rect[1])
            compatible_dc.SelectObject(data_bitmap)
            compatible_dc.BitBlt((0, 0), (self.rect[0], self.rect[1]), dc_obj, (0, 0), win32con.SRCCOPY)
            if save:
                data_bitmap.SaveBitmapFile(compatible_dc, 'ss.jpg')
            b = data_bitmap.GetBitmapBits(True)
            dc_obj.DeleteDC()
            compatible_dc.DeleteDC()
            win32gui.ReleaseDC(self.handle, window_dc)
            win32gui.DeleteObject(data_bitmap.GetHandle())
            self.current_img = [b, self.rect]
            self.current_npimg = np.frombuffer(self.current_img[0], dtype='uint8')
            self.current_npimg.shape = (self.current_img[1][1], self.current_img[1][0], 4)
        except pywintypes.error:
            pass
