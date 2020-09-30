"""
Clipboard Saver V0.26
By Steve Shambles Sept 2020
Windows only
Tested on Windows 7, 64bit.

Requirements:
cs2.ico (supplied) in current dir.
If not found then default sytem icon is used.

pip install Pillow
pip install pyperclip
pip install infi.systray
pip install pyautogui

V0.26:
-------
Added Save current screen option where a screenshot is taken,
named with a timestamp and saved.

Added Viewcontents of clipboard (text or img) using systems
default viewers.

Updated help
Added Donate me option
Added Contact author option

V0.26-1
Two tiny typos in help pop up corrected, doh!
"""
import ctypes
from datetime import datetime
import os
import time
import winsound
import webbrowser as web

from infi.systray import SysTrayIcon as stray
from PIL import ImageGrab
import pyautogui
import pyperclip

#  Check required folders exist in current dir.
#  If they don't then create them.
img_folder = 'clipboard_images'
check_img_folder = os.path.isdir(img_folder)
if not check_img_folder:
    os.makedirs(img_folder)

txt_folder = 'clipboard_texts'
check_txt_folder = os.path.isdir(txt_folder)
if not check_txt_folder:
    os.makedirs(txt_folder)

# Put copy of screen on clipboard. This solves a small bug
# where if the first thing you do is click save screen an error occurs.
pyautogui.press('printscreen')

def beep_sound(stray):
    """Beep sound. frequency, duration."""
    winsound.Beep(840, 100)

def clear_cb(stray):
    """Clear the clipboard and then make a beep sound to alert the user."""
    pyperclip.copy('')
    beep_sound(stray)

def save_cb_txt(stray):
    """If text found on clipboard, Save to uniquely named text file."""
    #  Grab clipboard contents.
    cb_txt = pyperclip.paste()

    if cb_txt:
        undr_ln = '-' *43+'\n'
        #  Properly readable date and time as the file name.
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

def save_curr_screen(stray):
     pyautogui.press('printscreen')
     save_cb_img(stray)

def save_clipb(stray):
    """When Save clipboard is selected call the two check and save functions."""
    save_cb_img(stray)
    save_cb_txt(stray)
    #  Pause is to make sure cannot save more than once a second
    #  or might interfere with unique filenames.
    time.sleep(1)

def view_clipb(stray):
    """View either text or image from clipboard using systems default viewers."""
    #  check for image on cb.
    img = ImageGrab.grabclipboard()
    if img:
        img.save('temp.jpg')
        web.open('temp.jpg')
        return

    # is it text on cb?
    cb_txt = ''
    cb_txt = pyperclip.paste()
    if cb_txt:
        with open ('temp.txt', 'w') as contents:
            contents.write(cb_txt)
        web.open('temp.txt')

def exit_prg(stray):
    """When quit is clicked, thread should close and icon in systray destroyed."""
    pass

def open_folder(stray):
    """Get current dir and open systems file browser to view contents."""
    cwd = os.getcwd()
    web.open(cwd)

def visit_blog(stray):
    """Open web and go to my Python blog site."""
    web.open('https://stevepython.wordpress.com/2020/03/16/how-to-create-a-systray-app')

def donate_me(stray):
    web.open("https://paypal.me/photocolourizer")

def contact_me(stray):
    """Open web and go to my blog contact page."""
    web.open('https://stevepython.wordpress.com/contact')

def cb_help(stray):
    """Popup box describing how to use the program."""
    ctypes.windll.user32.MessageBoxW(None,
                                     u'\nClipboard Saver V0.26.\n'
                                     'Freeware by Steve Shambles Sept 2020.\n\n'
                                     'Use the right click menu options to click'
                                     ' on "Save Clipboard"\nto save the current '
                                     'contents of the clipboard.\n'
                                     'Text and images are supported\n\n'
                                     'This can also be achieved by double clicking\n'
                                     'on the Clipboard Saver icon.\n\n'
                                     'If found, either Text or an image will\n'
                                     'be saved to the relevant folder.\n'
                                     'A beep will sound if something is saved.\n\n'
                                     'Selecting the "Save current screen" option will\n'
                                     'take a screenshot of your current screen and\n'
                                     'save it to the "clipboard_images" folder.\n'
                                     'A beep will again confirm if saved.\n\n'
                                     'If you click "Clear clipboard" a beep will\n'
                                     'also sound to confirm the action.\n\n'
                                     'Selecting "View Clipboard" will display either\n'
                                     'an image or text using your systems default\n'
                                     'viewer programs, assuming something is found\n\n'
                                     'You can view your saved data by clicking on the'
                                     '"Containing folder" item\n'
                                     'and looking in the "clipboard_text" and'
                                     '"clipboard_images" folders.\n\n'
                                     'If you find this program useful please\n'
                                     'consider making a small donation.\n\n',
                                     u'Clipboard Saver V0.26', 0)

def about_cbsaver(stray):
    """About program pop up."""
    ctypes.windll.user32.MessageBoxW(None,
                                     u'\nCB Saver V0.26.\n'
                                     'Freeware by Steve Shambles Sept 2020.\n\n'
                                     'CB Saver is a clipboard manager\n'
                                     'written in Python.\n\n'
                                     'Source code available free on my site and GitHub.\n\n'
                                     'If you find this program useful\n'
                                     'please consider making a small donation.\n\n',
                                     u'Clipboard Saver V0.26', 0)

hover_text = "Clipboard Saver V0.26"

#  Replace None with "your.ico" to use icons in the menus if you want.
menu_options = (('Save clipboard', None, save_clipb),
                ('Clear clipboard', None, clear_cb),
                ('Save current screen', None, save_curr_screen),
                ('View Clipboard', None, view_clipb),
                ('Open containing folder', None, open_folder),
                ('Options', None, (('Help', None, cb_help),
                                    ('About', None, about_cbsaver),
                                   ('Contact author', None, contact_me),
                                   ('Make a small donation to author', None, donate_me),
                                   ('Visit my blog', None, visit_blog),))
               )
#  Main program icon that will appear in systray. If not found system default icon used.
stray = stray("cs2.ico",
              hover_text, menu_options, on_quit=exit_prg,
              default_menu_index=0)
stray.start()
