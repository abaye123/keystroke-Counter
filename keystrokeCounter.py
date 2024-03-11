# keystroke Counter v0.2.0
# Date: 11/03/2024
# Email: ca@abaye.co
# GitHub: github.com/abaye123

import os
import keyboard
import pandas as pd
import threading
#import ctypes
from ctypes import create_unicode_buffer, windll, wintypes
import time
from datetime import datetime
import pytz


def get_desktop_folder():
    home_dir = os.path.expanduser("~")
    desktop_folder = os.path.join(home_dir, 'Desktop')
    return desktop_folder


def get_foreground_window_title():
    hwnd = windll.user32.GetForegroundWindow()
    length = windll.user32.GetWindowTextLengthW(hwnd) + 1
    title = create_unicode_buffer(length)
    windll.user32.GetWindowTextW(hwnd, title, length)
    return title.value

def record_typing(start_time, end_time, keystrokes, software_name, output_file):
    data = {'תחילת הכתיבה': [start_time],
            'סיום הכתיבה': [end_time],
            'מספר ההקשות': [keystrokes],
            'שם התוכנה': [software_name]}
    
    df = pd.DataFrame(data)
    
    #df['תחילת הכתיבה'] = pd.to_datetime(df['תחילת הכתיבה'], unit='s').dt.strftime('%d/%m/%Y %H:%M:%S')
    #df['סיום הכתיבה'] = pd.to_datetime(df['סיום הכתיבה'], unit='s').dt.strftime('%d/%m/%Y %H:%M:%S')

    df['תחילת הכתיבה'] = pd.to_datetime(df['תחילת הכתיבה'], unit='s')
    df['סיום הכתיבה'] = pd.to_datetime(df['סיום הכתיבה'], unit='s')
    
    israel_tz = pytz.timezone('Asia/Jerusalem')
    df['תחילת הכתיבה'] = df['תחילת הכתיבה'].dt.tz_localize('UTC').dt.tz_convert(israel_tz)
    df['סיום הכתיבה'] = df['סיום הכתיבה'].dt.tz_localize('UTC').dt.tz_convert(israel_tz)
    
    df['תחילת הכתיבה'] = df['תחילת הכתיבה'].dt.strftime('%d/%m/%Y %H:%M:%S')
    df['סיום הכתיבה'] = df['סיום הכתיבה'].dt.strftime('%d/%m/%Y %H:%M:%S')
    
    try:
        existing_data = pd.read_excel(output_file, engine='openpyxl')
        df = pd.concat([existing_data, df], ignore_index=True)
    except FileNotFoundError:
        pass
     
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Data recorded for {software_name}")

def keyboard_listener(output_file):
    start_time = time.time()
    software_name = get_foreground_window_title()
    end_time = start_time
    keystrokes = 0
    
    def on_key_event(event):
        nonlocal start_time, end_time, keystrokes, software_name
        if event.event_type == keyboard.KEY_DOWN:
            keystrokes += 1
            end_time = time.time()
    
    keyboard.hook(on_key_event)
    
    try:
        while True:
            time.sleep(1)
            current_software_name = get_foreground_window_title()
            if current_software_name != software_name:
                record_typing(start_time, end_time, keystrokes, software_name, output_file)
                start_time = time.time()
                software_name = current_software_name
                keystrokes = 0
    except KeyboardInterrupt:
        pass
    finally:
        keyboard.unhook_all()
        record_typing(start_time, end_time, keystrokes, software_name, output_file)

def main():
    current_timestamp = time.time()
    current_datetime = datetime.utcfromtimestamp(current_timestamp)
    current_date = current_datetime.strftime('%d-%m-%Y')

    desktop_path = get_desktop_folder()
    folder_path = f'{desktop_path}\מונה הקשות'

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Folder '{folder_path}' created.")

    output_file_path = f'{folder_path}\{current_date}.xlsx'
    print("output path: " + output_file_path)
    print("Recording keystrokes...")
    
    listener_thread = threading.Thread(target=keyboard_listener, args=(output_file_path,))
    listener_thread.start()
    
    try:
        listener_thread.join()
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
