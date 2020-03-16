"""
Clipboard Saver V0.24
By Steve Shambles March 2020
https://stevepython.wordpress.com/2020/03/16/how-to-create-a-systray-app
Windows only
Tested on Windows 7, 64bit.

Requirements:
cs2.ico (supplied) in current dir.
If not found then default sytem icon is used.

pip install Pillow
pip install pyperclip
pip install infi.systray
"""
import ctypes
from datetime import datetime
import os
import time
import winsound
import webbrowser

from infi.systray import SysTrayIcon as stray
from PIL import ImageGrab
import pyperclip

# Check required folders exist in current dir.
# If they don't then create them.
img_folder = 'clipboard_images'
check_img_folder = os.path.isdir(img_folder)
if not check_img_folder:
    os.makedirs(img_folder)

txt_folder = 'clipboard_texts'
check_txt_folder = os.path.isdir(txt_folder)
if not check_txt_folder:
    os.makedirs(txt_folder)

def beep_sound(stray):
    """Beep sound. frequency, duration."""
    winsound.Beep(840, 100)

def clear_cb(stray):
    """Clear the clipboard and then make a beep sound to alert the user."""
    pyperclip.copy('')
    beep_sound(stray)

def save_cb_txt(stray):
    """If text found on clipboard, Save to uniquely named text file."""
    # Grab clipboard contents.
    cb_txt = pyperclip.paste()

    if cb_txt:
        undr_ln = '-' *43+'\n'
        # I wanted a properly readable date and time as the file name.
        time_stamp = (datetime.now().strftime
                      (r'%d'+('-')+'%b'+('-')+'%Y'+('-')+'%H'+('.')
                       +'%M'+('-')+'%S'+'s'))
        file_name = time_stamp+'.txt'
        folder = r'clipboard_texts/'

        with open(folder+str(file_name), 'w') as contents:
            contents.write('Clipboard text found: '+str(time_stamp)+'\n')
            contents.write(undr_ln)
            contents.write(cb_txt)
            beep_sound(stray)

def save_cb_img(stray):
    """If image found, Save to uniquely named jpg file."""
    img = ImageGrab.grabclipboard()

    if img:
        img_file_name = (datetime.now().strftime
                         (r'%d'+('-')+'%b'+('-')+'%Y'+('-')+'%H'+('.')
                          +'%M'+('-')+'%S'+'s'))+'.jpg'

        cb_img_folder = r'clipboard_images/'
        img.save(cb_img_folder+img_file_name)
        beep_sound(stray)

def save_clipb(stray):
    """When Save clipboard is selected call the two check and save functions."""
    save_cb_img(stray)
    save_cb_txt(stray)
    # Pause is to make sure cannot save more than once a second
    # or might interfere with unique filenames.
    time.sleep(1)

def exit_prg(stray):
    """When quit is clicked, thread should close and icon in systray destroyed."""
    pass

def open_folder(stray):
    """Get current dir and open systems file browser to view contents."""
    cwd = os.getcwd()
    webbrowser.open(cwd)

def visit_blog(stray):
    """Open webbrowser and go to my Python blog site."""
    webbrowser.open('https://stevepython.wordpress.com/2020/03/16/how-to-create-a-systray-app')

def cb_help(stray):
    """Popup box describing how to use the program."""
    ctypes.windll.user32.MessageBoxW(None,
                                     u'\nClipboard Saver.\n'
                                     'Freeware by Steve Shambles March 2020.\n\n'
                                     'Use the right click menu options to click'
                                     ' on "Save Clipboard"\nto save the current '
                                     'contents of the clipboard.\n\n'
                                     'This can also be achieved by double clicking\n'
                                     'on the Clipboard Saver icon.\n\n'
                                     'If found, either Text or an image will\n'
                                     'be saved to the relevant folder.\n'
                                     'A beep will sound if something is saved.\n\n'
                                     'If you click "Clear clipboard" a beep will\n'
                                     'also sound to confirm the action.\n\n'
                                     'You can view your saved data by clicking on the'
                                     '"Containing folder" item\n'
                                     'and looking in the "clipboard_text" and'
                                     '"clipboard_images" folders.\n\n',
                                     u'Clipboard Saver V0.24', 0)

# Program name.
hover_text = "Clipboard Saver V0.24"

# Replace None with "your.ico" to use icons in the menus if you want.
menu_options = (('Save clipboard', None, save_clipb),
                ('Clear clipboard', None, clear_cb),
                ('Open containing folder', None, open_folder),
                ('Options', None, (('Help', None, cb_help),
                                   ('Visit my blog', None, visit_blog),))
               )
# Main program icon that will appear in systray. If not found system default icon used.
stray = stray("cs2.ico",
              hover_text, menu_options, on_quit=exit_prg,
              default_menu_index=0)
stray.start()
