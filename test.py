import win32gui
import time
time.sleep(5)
print(win32gui.GetWindowText(win32gui.GetForegroundWindow()))

