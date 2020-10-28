import time
import sys

import psutil
import win32con

import detection
import winapihelper
from screenshot import Screenshot


def is_valid_coords(coords):
    return coords.__len__() > 0


class MainLoop:
    def __init__(self, order_number):
        self.pid = 0
        self.screenshot = None
        self.handle = [0, 0]
        self.finished = False
        self.successful = False
        self.order_number = order_number

    def update_screenshot(self):
        # checks if the window is minized and restores it, otherwise, the screenshot returns black
        winapihelper.restore_window_if_minimized(self.handle[0])
        self.screenshot = Screenshot(self.handle[0])
        self.screenshot.get_image(True)

    def check_if_connection_is_alive(self):
        self.pid = winapihelper.get_rdp_pid()
        return not self.pid == 0

    def close_popups(self):
        # tries to find any popups open, and then closes it
        close_button = detection.find_popup_close_button(self.screenshot.current_npimg)
        if close_button[0].__len__() > 0:
            winapihelper.mouse_click(self.handle[1], close_button)
            time.sleep(1)
            self.update_screenshot()
            self.close_popups()

    def go_to_main_window(self):
        # tries to go back to the main window
        if not is_valid_coords(detection.check_if_in_main_window(self.screenshot.current_npimg)[0]):
            main_window_taskbar = \
                detection.find_main_window_in_taskbar(self.screenshot.current_npimg)
            winapihelper.mouse_click(self.handle[1], main_window_taskbar)
            time.sleep(1)
            self.update_screenshot()
            self.go_to_main_window()

    def open_print_order_window(self):
        # look for the post sell button and clicks it
        button_coords = detection.find_post_sell_button(self.screenshot.current_npimg)
        if is_valid_coords(button_coords[0]):
            winapihelper.mouse_click(self.handle[1], button_coords)
            time.sleep(1)
            self.update_screenshot()

            order_button_coords = detection.find_orders_button(self.screenshot.current_npimg)
            if is_valid_coords(order_button_coords[0]):
                winapihelper.mouse_click(self.handle[1], order_button_coords)
                time.sleep(1)
                self.update_screenshot()
            else:
                self.open_print_order_window()

            stock_manage_button_coords = detection.find_stock_manage_button(self.screenshot.current_npimg)
            if stock_manage_button_coords[0].__len__() > 0:
                winapihelper.mouse_click(self.handle[1], stock_manage_button_coords)
                time.sleep(1)
                self.update_screenshot()
            else:
                self.open_print_order_window()
        else:
            self.open_print_order_window()

    def print_order_routine(self):
        self.close_popups()
        # checks if the order window is already open
        if detection.check_if_order_window_is_open(self.screenshot.current_npimg)[0].__len__() > 0:
            search_invoice_button = detection.find_search_invoice_button(self.screenshot.current_npimg)
            if is_valid_coords(search_invoice_button[0]):
                # we want the click to be slightly to the left, since we want to edit the
                # textbox, not click the button
                winapihelper.mouse_click(self.handle[1], search_invoice_button, (-40, 0))
                winapihelper.mouse_click(self.handle[1], search_invoice_button, (-40, 0))

                # write order number to field
                winapihelper.write_text(self.handle[1], str(self.order_number))

                # press enter, to load the order and close the popup
                winapihelper.keyboard_click(self.handle[1], win32con.VK_RETURN)
                time.sleep(.25)
                winapihelper.keyboard_click(self.handle[1], win32con.VK_RETURN)

                self.update_screenshot()

                print_button = detection.find_print_button(self.screenshot.current_npimg)
                if is_valid_coords(print_button[0]):
                    winapihelper.mouse_click(self.handle[1], print_button)
                    time.sleep(1)
                    self.update_screenshot()
                else:
                    self.print_order_routine()

                print_order_button = detection.find_print_order_button(self.screenshot.current_npimg)
                if is_valid_coords(print_order_button[0]):
                    winapihelper.mouse_click(self.handle[1], print_order_button)
                    time.sleep(1)
                    self.update_screenshot()
                else:
                    self.print_order_routine()

                ok_button = detection.find_ok_button(self.screenshot.current_npimg)
                if is_valid_coords(ok_button[0]):
                    winapihelper.mouse_click(self.handle[1], ok_button)
                    time.sleep(3)
                    self.update_screenshot()
                else:
                    self.print_order_routine()

                # checks if the systems tells us that theres nothing to be printed
                check_if_printable = detection.check_if_cannot_print(self.screenshot.current_npimg)
                if is_valid_coords(check_if_printable[0]):
                    self.close_popups()
                    self.finished = True
                    self.successful = True
                    return 1

                # press tab, to unselect the file type combobox
                winapihelper.keyboard_click(self.handle[1], win32con.VK_TAB)
                time.sleep(.25)

                save_to_file_button = detection.save_to_file_button(self.screenshot.current_npimg)
                if is_valid_coords(save_to_file_button[0]):
                    winapihelper.mouse_click(self.handle[1], save_to_file_button)
                    self.close_popups()
                    self.finished = True
                    self.successful = True
                else:
                    self.print_order_routine()
            else:
                self.print_order_routine()

        else:
            self.go_to_main_window()
            self.open_print_order_window()
            self.print_order_routine()

    def run(self):
        while not self.finished:
            # looks for the rdp connection
            self.check_if_connection_is_alive()
            # gets the handle based on the PID
            self.handle = winapihelper.get_handle(self.pid)
            if self.pid == 0:
                try:
                    winapihelper.kill_rdp_error_windows()
                    if self.handle[0] != 0:
                        winapihelper.kill_inactive_windows()
                    winapihelper.open_rdp()
                    time.sleep(5)
                except psutil.AccessDenied:
                    time.sleep(1)
            else:
                if self.handle[0] != 0:
                    self.update_screenshot()
                    is_system_open = detection.check_if_system_is_open(self.screenshot.current_npimg)
                    if not is_system_open:
                        time.sleep(5)
                        if not self.open_system():
                            break

                        self.update_screenshot()

                    self.print_order_routine()
                    self.finished = True
                    self.successful = True
            time.sleep(1)
        print(self.successful)

    def open_system(self):
        coords = detection.find_program_exe(self.screenshot.current_npimg)
        if is_valid_coords(coords[0]):
            winapihelper.mouse_click(self.handle[1], coords)
            winapihelper.mouse_click(self.handle[1], coords)
            for tries in range(0, 60):
                if not detection.check_if_system_is_open(self.screenshot.current_npimg):
                    print('waiting for the program to start.')
                    time.sleep(1)
                    self.update_screenshot()
                elif tries == 59:
                    self.finished = True
                    return False
                else:
                    time.sleep(3)
                    break
        return True

    def only_open_system(self):
        while self.pid == 0:
            self.check_if_connection_is_alive()
            # gets the handle based on the PID
            self.handle = winapihelper.get_handle(self.pid)
            if self.pid == 0:
                try:
                    winapihelper.kill_rdp_error_windows()
                    if self.handle[0] != 0:
                        winapihelper.kill_inactive_windows()
                    winapihelper.open_rdp()
                    time.sleep(5)
                except psutil.AccessDenied:
                    time.sleep(1)
            time.sleep(1)


if __name__ == '__main__':
    if sys.argv.__len__() == 2:
        routine = MainLoop(sys.argv[1])
        if sys.argv[1] == 'open':
            routine.only_open_system()
        else:
            print('imprimindo pedido ' + sys.argv[1])
            routine.run()
    else:
        print('numero de pedido invalido')
