import time, util, datetime
import win32gui
import win32api
import win32con
import os
import sys
from PIL import ImageGrab
from pynput.keyboard import Controller
from datetime import datetime
import pytesseract
from PIL import Image


DEBUG = False
EMPTY_STRING = ''
SKIP_FIELD_VALUE = "<skip>"
TOP_LEFT_FIELD = "{"

import cv2
import numpy as np

last_tab = -1
failed_count = 0

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
    util.write_file(test_result_folder + filename + '.txt', msg, test_result_folder)

def compare_result_text(filename: str, message_text: str, test_result_folder: str, test_case: str, claim: util.ClaimsData, field_map):
    global failed_count
    # Open the image file
    image = Image.open(filename)

    # Perform OCR using pytesseract
    text = pytesseract.image_to_string(image)

    # Print the recognized text
    if DEBUG:
        print(text)

        print(message_text in text)

    result = "Passed" if message_text in text else "Failed"

    text = "Expected: " + message_text + "\n Results stored in " + filename + '.' + result + '.txt'

    util.write_file(filename + '.' + result + '.txt', result)

    if result == "Failed":
        failed_count = failed_count + 1
        col = find_key(field_map, test_case)
        index = claim.data_header.fields.index(col)
        util.append_file(test_result_folder + "_failed_test_summary.txt", "test number: " + claim.data[index] + " - " + text + "\n")

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

    if DEBUG:
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

def send_keystrokes(keys, whnd: int, pause = True):
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
        elif key == "rev_tab":
            # For uppercase letters, send the Shift key first
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
            win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif key == '*':
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(ord('8'), 0, 0, 0)
            #win32api.keybd_event(ord('8'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif key == "-" or key == TOP_LEFT_FIELD or key == ".":
            keyboard = Controller()
            keyboard.press(key)
            keyboard.release(key)
        else:
            # For all other keys, send the key directly
            win32api.keybd_event(ord(key), 0, 0, 0)
            win32api.keybd_event(ord(key), 0, win32con.KEYEVENTF_KEYUP, 0)
        if pause:
            time.sleep(.1)

    return
    
def start_top_left(whnd: int):
    global last_tab
    #send_keystrokes(TOP_LEFT_FIELD, whnd)
    for x in range(0,last_tab + 1):
        send_reverse_tab(whnd)
    #time.sleep(.5)

    return

def send_enter(whnd: int):
    # Send the Enter key press to the window
    _send_controls_keys(win32con.VK_RETURN, whnd)

    return

def send_tab(whnd: int):
    # Send the F1 key press to the window
    _send_controls_keys(win32con.VK_TAB, whnd)

    return

def send_reverse_tab(whnd: int):
    # Send the F1 key press to the window
    send_keystrokes(["rev_tab"], whnd, False)

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

def send_f12(whnd: int):
    # Send the F12 key press to the window
    _send_controls_keys(win32con.VK_F12, whnd)

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

def perform_search(claim: util.ClaimsData, field_map, whnd: int, system_index: int, tran_index: int):
    result_pos = -1
    result_tab_pos = -1
    for map in field_map['fields']:
        for field_type in map:
            if field_type == "search":
                for field_type_field in map[field_type]:
                    if field_type_field["name"] == "results":
                        result_pos = claim.data_header.fields.index(field_type_field["name"])
                        result_tab_pos = field_type_field["tab_pos"]
                    elif field_type_field["name"] == "submit":
                        submit_keys = claim.data[claim.data_header.fields.index(field_type_field["name"])].split("|")
                    elif field_type_field["name"] == "select":
                        select_keys = claim.data[claim.data_header.fields.index(field_type_field["name"])].split("|")

    send_data("search", field_map, ["results","submit","select"], claim, whnd)

    for submit_key in submit_keys:
        if submit_key.strip().lower() == "<f4>":
            send_f4(whnd)
    time.sleep(1)
    
    for x in range(0,result_tab_pos):
        send_tab(whnd)
    send_keystrokes(claim.data[result_pos], whnd)
    for select_key in select_keys:
        if select_key.strip().lower() == "<enter>":
            send_enter(whnd)

    return

def fill_out_page(page: str, claim: util.ClaimsData, field_map, whnd):
    global last_tab
    last_tab = -1
    send_data(page, field_map, [], claim, whnd)
    send_enter(whnd)

    return

def send_data(field_group: str, field_map, ignore_fields, claim: util.ClaimsData, whnd: int):
    global last_tab
    for map in field_map['fields']:
        for field_type in map:
            if field_type == field_group:
                for field_type_field in map[field_type]:
                    data_col = -1
                    tab_pos = -1
                    length = 0
                    for ftk, ftv in field_type_field.items():
                        if ftk == "name":
                            if ftv in ignore_fields:
                                continue
                            if ftv in claim.data_header.fields:
                                data_col = claim.data_header.fields.index(ftv)
                        elif ftk == "tab_pos":
                            tab_pos = ftv
                        elif ftk == "length":
                            length = int(ftv)
                
                    if data_col > -1:
                        if claim.data[data_col].strip().lower() != SKIP_FIELD_VALUE:
                            start_top_left(whnd)
                            for x in range(0,tab_pos):
                                send_tab(whnd)
                            send_keystrokes(claim.data[data_col], whnd)
                            l = len(claim.data[data_col])

                            if l < length:
                                send_keystrokes(util.pad_char(length - l, " "), whnd)
                            last_tab = tab_pos
                            time.sleep(.5)

    return

def create_test_result_folder():
    # Get the current date and time
    now = datetime.now()

    # Format the date and time as a string
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")

    # Create a new folder with the timestamp as its name
    new_folder = os.path.join(os.getcwd(), "results/", timestamp)
    os.mkdir(new_folder)

    return new_folder + "/"

def main(system: str, tran: str):
    global last_tab
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

        system_identifier = "system"
        transaction_identifier = "transaction_type"
        compare_identifier = "compare"
        final_action_identifier = "final_action"
        change_request_identifier = "change_request"
        business_requirement_identifier = "business_requirement"
        user_story_identifier = "user_story"
        test_case_identifier = "test_case"

        filename = system.lower().strip() + "_" + tran.lower().strip() + '_field_map.json'
        field_map = util.get_field_map(filename)

        system_col = find_key(field_map, system_identifier)
        compare_col = find_key(field_map, compare_identifier)
        tran_col = find_key(field_map, transaction_identifier)

        claims = util.read_claims_data(tran.lower().strip())

        now = datetime.now()
        print(now.strftime("%d/%m/%Y %H:%M:%S"))
        print("Total tests: " + str(len(claims)))
        count = 0

        for claim in claims:
            last_tab = -1
            system_index = claim.data_header.fields.index(system_col)
            compare_index = claim.data_header.fields.index(compare_col)
            tran_index = claim.data_header.fields.index(tran_col)
            launch_system(claim.data[system_index], whnd)

            perform_search(claim, field_map, whnd, system_index, tran_index)

            fill_out_page('page 1', claim, field_map, whnd)
            fill_out_page('page 2', claim, field_map, whnd)
            fill_out_page('page 3', claim, field_map, whnd)
            fill_out_page('page 4', claim, field_map, whnd)
            fill_out_page('page 5', claim, field_map, whnd)
            fill_out_page('page 6', claim, field_map, whnd)
            fill_out_page('page 7', claim, field_map, whnd)
            fill_out_page('page 8', claim, field_map, whnd)
            fill_out_page('page 9', claim, field_map, whnd)
            fill_out_page('page 10', claim, field_map, whnd)

            final_action = find_key(field_map, final_action_identifier)

            final_action_col = claim.data_header.fields.index(final_action)

            if claim.data[final_action_col] == "F4":
                send_f4(whnd)            
            elif claim.data[final_action_col] == "F5":
                send_f5(whnd)
            elif claim.data[final_action_col] == "F12":
                send_f12(whnd)

            time.sleep(1)
            file = take_screenshot(whnd, [change_request_identifier, business_requirement_identifier, user_story_identifier, test_case_identifier], claim, field_map, test_result_folder)
            time.sleep(1)
            compare_result_text(test_result_folder + file, claim.data[compare_index], test_result_folder, test_case_identifier, claim, field_map)
            send_f3(whnd)

            count = count + 1
            print("Finished test: " + str(count) + " of " + str(len(claims)))

        print("Total failed tests: " + str(failed_count))
        print("Tests completed")
        now = datetime.now()
        print(now.strftime("%d/%m/%Y %H:%M:%S"))
    else:
        print("Window not found")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        main("HITF", "HUOP")
    else:
        main(sys.argv[1], sys.argv[2])