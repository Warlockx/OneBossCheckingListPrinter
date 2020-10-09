import cv2
import numpy as np
from PIL import ImageColor


def load_template(file):
    return cv2.imread('templates/' + file, 0)


def get_grayscale(source):
    return cv2.cvtColor(source, cv2.COLOR_BGR2GRAY)


def draw_debug_shape(source, template, coords):
    w, h, = template.shape[::-1]
    for c in coords:
        cv2.rectangle(source, c, (c[0] + w, c[1] + h), ImageColor.getrgb('red'), 2)
    cv2.imwrite('DEBUG.jpg', source)


def get_coords(source, template, template_name, debug=False, threshold=.9):
    grayscale = get_grayscale(source)
    matches = cv2.matchTemplate(grayscale, template, cv2.TM_CCOEFF_NORMED)
    good_matches = np.where(matches >= threshold)
    coords = []
    for point in zip(*good_matches[::-1]):
        coords.append((point[0], point[1]))

    if debug:
        if coords.__len__() > 0:
            print(template_name + ' found')
            draw_debug_shape(source, template, coords)
        else:
            print(template_name + ' not found')
    return coords, template


def check_if_system_is_open(source, debug=False):
    template = load_template('window_open_taskbar.png')
    coords = get_coords(source, template, 'window_open_taskbar', debug, .4)
    return coords.__len__() > 0


def find_program_exe(source, debug=False):
    template = load_template('program_exe.png')
    return get_coords(source, template, 'program_exe', debug, .8)


def check_if_order_window_is_open(source, debug=False):
    template = load_template('order_window_detect.png')
    return get_coords(source, template, 'order_window_detect', debug)


def check_if_post_sell_button_is_active(source, debug=False):
    template = load_template('post_sell_button_active.png')
    return get_coords(source, template, 'post_sell_button_active', debug, 1)


def find_popup_close_button(source, debug=False):
    template = load_template('pop_up_close_button.png')
    return get_coords(source, template, 'pop_up_close_button')


def find_post_sell_button(source, debug=False):
    template = load_template('post_sell_button.png')
    return get_coords(source, template, 'post_sell_button', debug, .9)


def find_main_window_in_taskbar(source, debug=False):
    template = load_template('window_open_taskbar.png')
    return get_coords(source, template, 'window_open_taskbar', debug, .4)


def check_if_in_main_window(source, debug=False):
    template = load_template('main_page_indicator.png')
    return get_coords(source, template, 'main_page_indicator', debug, .9)


def find_orders_button(source, debug=False):
    template = load_template('orders_button.png')
    template_active = load_template('orders_button_active.png')
    inactive_coords = get_coords(source, template, 'orders_button', debug)
    active_coords = get_coords(source, template_active, 'actived_order_button', debug)
    return max(inactive_coords, active_coords), template


def find_stock_manage_button(source, debug=False):
    template = load_template('stock_manage_button.png')
    return get_coords(source, template, 'stock_manage_button', debug)


def find_search_invoice_button(source, debug=False):
    template = load_template('search_invoice_button.png')
    return get_coords(source, template, 'search_invoice_button', debug)


def find_print_button(source, debug=False):
    template = load_template('print_button.png')
    return get_coords(source, template, 'print_button', debug)


def find_print_order_button(source, debug=False):
    template = load_template('print_order_button.png')
    return get_coords(source, template, 'print_order_button', debug)


def find_ok_button(source, debug=False):
    template = load_template('ok_button.png')
    return get_coords(source, template, 'ok_button', debug)


def save_to_file_button(source, debug=False):
    template = load_template('save_to_file_button.png')
    return get_coords(source, template, 'save_to_file_button', debug)



