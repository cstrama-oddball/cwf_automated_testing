import time, util, datetime
import win32gui
import win32api
import win32con
import os
from PIL import ImageGrab
from pynput.keyboard import Controller
import pytesseract
from PIL import Image


DEBUG = False
EMPTY_STRING = ''
SKIP_FIELD_VALUE = "<skip>"
TOP_LEFT_FIELD = "{"

import cv2
import numpy as np

def compare_result(search_for: str, filename: str, test_result_folder: str):
    if DEBUG:
        print(search_for)
        print(filename)
    # Read in the search image and the template image
    search_image = cv2.imread(search_for)
    search_in = cv2.imread(test_result_folder + filename)

    # Convert the images to grayscale
    search_image_gray = cv2.cvtColor(search_image, cv2.COLOR_BGR2GRAY)
    template_image_gray = cv2.cvtColor(search_in, cv2.COLOR_BGR2GRAY)

    # Compute the matching scores
    result = cv2.matchTemplate(search_image_gray, template_image_gray, cv2.TM_CCOEFF_NORMED)

    # Find the location with the highest score
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # Set a threshold for the matching score
    threshold = 0.9

    # Check if the highest score is above the threshold
    if max_val > threshold:
        msg = 'The template image is present in the search image'        
    else:
        msg = 'The template image is not present in the search image'

    print(msg)
    util.write_file(test_result_folder + filename + '.txt', msg)

def compare_result_text(filename: str, message_text: str):
    # Open the image file
    image = Image.open(filename)

    # Perform OCR using pytesseract
    text = pytesseract.image_to_string(image)

    # Print the recognized text
    if DEBUG:
        print(text)

        print(message_text in text)

    result = "Passed" if message_text in text else "Failed"

    util.write_file(filename + '.' + result + '.txt', result)

def take_screenshot(whnd: int, filename_fields, claim: util.ClaimsData, field_map, test_result_folder: str): 
    filename = EMPTY_STRING

    for filename_field in filename_fields:
        system_col = find_key(field_map, filename_field)
        system_index = claim.data_header.fields.index(system_col)
        filename = filename + system_col + "_" + claim.data[system_index] + "_"

    filename = filename[0:len(filename) - 1]
    # Get the dimensions of the current window
    left, top, right, bottom = win32gui.GetWindowRect(whnd)
    if DEBUG:
        print('left ' + str(left))
        print('top ' + str(top))
        print('right ' + str(right))
        print('bottom ' + str(bottom))
    top = top + 20
    left = left + 31
    right = right + 615
    bottom = bottom + 328
    width = right - left
    height = bottom - top

    if DEBUG:
        print(str(width))
        print(str(height))

    # Take a screenshot of the current window
    screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))

    # Save the screenshot to a file
    filename = filename + '.png'
    screenshot.save(test_result_folder + filename)

    # Print the path of the saved file to the console
    print('Screenshot saved to', os.path.abspath(filename))

    return filename

def resize_window(whnd):
    new_width = 1215
    new_height = 640
    win32gui.SetWindowPos(whnd, None, 0, 0, new_width, new_height, win32con.SWP_NOMOVE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)

    return

def move_window(whnd):
    new_left = 35
    new_top = 42
    # Move the window
    win32gui.SetWindowPos(whnd, None, new_left, new_top, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)

    return

def send_keystrokes(keys, whnd: int):
    win32gui.SetForegroundWindow(whnd) 
    time.sleep(.5)
    # Send the keystrokes to the window
    for key in keys:
        if key.isupper():
            # For uppercase letters, send the Shift key first
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(ord(key), 0, 0, 0)
            win32api.keybd_event(ord(key), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif key == '*':
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(ord('8'), 0, 0, 0)
            #win32api.keybd_event(ord('8'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif key == "-" or key == TOP_LEFT_FIELD:
            keyboard = Controller()
            keyboard.press(key)
            keyboard.release(key)
        else:
            # For all other keys, send the key directly
            win32api.keybd_event(ord(key), 0, 0, 0)
            win32api.keybd_event(ord(key), 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(.1)

    return
    
def start_top_left(whnd: int):
    send_keystrokes(TOP_LEFT_FIELD, whnd)
    time.sleep(.5)

    return

def send_enter(whnd: int):
    # Send the Enter key press to the window
    _send_controls_keys(win32con.VK_RETURN, whnd)

    return

def send_tab(whnd: int):
    # Send the F1 key press to the window
    _send_controls_keys(win32con.VK_TAB, whnd)

    return

def send_f1(whnd: int):
    # Send the F1 key press to the window
    _send_controls_keys(win32con.VK_F1, whnd)

    return

def send_f3(whnd: int):
    # Send the F3 key press to the window
    _send_controls_keys(win32con.VK_F3, whnd)

    return

def send_f4(whnd: int):
    # Send the F4 key press to the window
    _send_controls_keys(win32con.VK_F4, whnd)

    return

def send_f5(whnd: int):
    # Send the F5 key press to the window
    _send_controls_keys(win32con.VK_F5, whnd)

    return

def send_f9(whnd: int):
    # Send the F4 key press to the window
    _send_controls_keys(win32con.VK_F9, whnd)

    return

def _send_controls_keys(key: int, whnd: int):
    win32gui.SetForegroundWindow(whnd) 
    time.sleep(.1)
    win32api.keybd_event(key, 0, 0, 0)
    win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)

    return

def launch_system(system: str, whnd: int):
    send_keystrokes(system, whnd)
    send_enter(whnd)
    time.sleep(2)

    return

def find_key(field_map, key: str):
    for map in field_map['fields']:
        for field_type in map:
            for field in map[field_type]:
                for k,v in field.items():
                    if k == key:
                        return v
                    
    return EMPTY_STRING

def search_for_claim(claim: util.ClaimsData, field_map, whnd: int, system_index: int):
    result_pos = -1
    result_tab_pos = -1
    for map in field_map['fields']:
        for field_type in map:
            if field_type == "search":
                for field_type_field in map[field_type]:
                    for ftk, ftv in field_type_field.items():
                        if ftk == "name":
                            if ftv == "results":
                                result_pos = claim.data_header.fields.index(ftv)
                        elif ftk == "tab_pos":
                            result_tab_pos = ftv

    send_data("search", field_map, ["results"], claim, whnd, system_index)

    send_f4(whnd)
    time.sleep(1)
    
    for x in range(0,result_tab_pos):
        send_tab(whnd)
    send_keystrokes(claim.data[result_pos], whnd)
    send_enter(whnd)

    return

def fill_out_page(page: str, claim: util.ClaimsData, field_map, whnd, system_index: int):
    send_data(page, field_map, [], claim, whnd, system_index)
    send_enter(whnd)

    return

def send_data(field_group: str, field_map, ignore_fields, claim: util.ClaimsData, whnd: int, system_index: int):
    for map in field_map['fields']:
        for field_type in map:
            if field_type == field_group:
                for field_type_field in map[field_type]:
                    data_col = -1
                    tab_pos = -1
                    same_system = False
                    for ftk, ftv in field_type_field.items():
                        if ftk == "name":
                            if ftv in ignore_fields:
                                continue
                            if ftv in claim.data_header.fields:
                                data_col = claim.data_header.fields.index(ftv)
                        elif ftk == "tab_pos":
                            tab_pos = ftv
                        elif ftk == "system":
                            same_system = ftv == claim.data[system_index]
                
                    if data_col > -1 and same_system:
                        if claim.data[data_col].strip().lower() != SKIP_FIELD_VALUE:
                            start_top_left(whnd)
                            for x in range(0,tab_pos):
                                send_tab(whnd)
                            send_keystrokes(claim.data[data_col], whnd)
                            time.sleep(.5)

    return

def create_test_result_folder():
    # Get the current date and time
    now = datetime.datetime.now()

    # Format the date and time as a string
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")

    # Create a new folder with the timestamp as its name
    new_folder = os.path.join(os.getcwd(), "results/", timestamp)
    os.mkdir(new_folder)

    return new_folder + "/"

#TODO: create a better way to store the window titles
win2find = 'Session A - [24 x 80]'

whnd = win32gui.FindWindowEx(None, None, None, win2find)
if not (whnd == 0):
    win32gui.SetForegroundWindow(whnd) 
    time.sleep(.5)
    resize_window(whnd)
    time.sleep(.5)
    move_window(whnd)
    time.sleep(.5)

    test_result_folder = create_test_result_folder()

    filename = 'field_map.json'
    field_map = util.get_field_map(filename)

    system_identifier = "system"
    compare_identifier = "compare"
    final_action_identifier = "final_action"
    change_request_identifier = "change_request"
    business_requirement_identifier = "business_requirement"
    user_story_identifier = "user_story"
    test_case_identifier = "test_case"

    system_col = find_key(field_map, system_identifier)
    compare_col = find_key(field_map, compare_identifier)

    claims = util.read_claims_data()

    for claim in claims:
        system_index = claim.data_header.fields.index(system_col)
        compare_index = claim.data_header.fields.index(compare_col)
        launch_system(claim.data[system_index], whnd)

        search_for_claim(claim, field_map, whnd, system_index)

        fill_out_page('page 1', claim, field_map, whnd, system_index)
        fill_out_page('page 2', claim, field_map, whnd, system_index)

        final_action = find_key(field_map, final_action_identifier)

        final_action_col = claim.data_header.fields.index(final_action)

        if claim.data[final_action_col] == "F4":
            send_f4(whnd)            
        elif claim.data[final_action_col] == "F5":
            send_f5(whnd)

        time.sleep(1)
        file = take_screenshot(whnd, [change_request_identifier, business_requirement_identifier, user_story_identifier, test_case_identifier], claim, field_map, test_result_folder)
        time.sleep(1)
        compare_result_text(test_result_folder + file, claim.data[compare_index])
        send_f3(whnd)
else:
    print("Window not found")