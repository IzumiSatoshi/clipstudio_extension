import subprocess
import time
import winsound
import win32gui
import pyautogui
import keyboard
import psutil
import pandas as pd

# 呼び出す時に使うホットキー
create_sub_window_hotkey = 'alt+a' #新規ウィンドウの作成
set_sub_window_rect_hotkey ='alt+p' #ウィンドウの位置を記憶
# クリスタ側のホットキー
new_window_hotkey = 'ctrl+shift+alt+p' #新規ウィンドウの作成
resize_area_hotkey = 'ctrl+shift+alt+o' #キャンバス全体を表示

step_px = 10
# アクティブなタブを探すときの、最初のピクセル
start_pixel = (0,60)
sub_window_rect = (100,100,1000,1000)
def main():
    keyboard.add_hotkey(create_sub_window_hotkey, create_sub_window)
    keyboard.add_hotkey(set_sub_window_rect_hotkey, set_sub_window_rect)
    global sub_window_rect
    try:
        sub_window_rect = pd.read_pickle('sub_window_rect.pkl')
    except FileNotFoundError:
        pass

    process = activate_clipstudio()
    process.wait()

def activate_clipstudio():
    exefile = r"C:\Program Files\CELSYS\CLIP STUDIO 1.5\CLIP STUDIO PAINT\CLIPStudioPaint.exe"

    # クリスタが起動しているかどうかを確認
    for proc in psutil.process_iter():
        try:
            if proc.exe() == exefile:
                print('クリップスタジオペイントは、既に起動済みです')
                return proc
        except psutil.AccessDenied:
            pass

    proc = subprocess.Popen(exefile)
    print('クリップスタジオペイントを起動します')
    return proc

def set_sub_window_rect():
    # ウィンドウの位置を保存する

    global sub_window_rect
    hwnd = win32gui.GetForegroundWindow()
    sub_window_rect = win32gui.GetWindowRect(hwnd)
    pd.to_pickle(sub_window_rect, "sub_window_rect.pkl")
    print("ウィンドウの位置を記憶しました " + str(sub_window_rect))

def create_sub_window():
    win_size = pyautogui.size()
    # いったんセンターにドラッグ
    center_pix = (900,500)
    keyboard.send(new_window_hotkey)
    time.sleep(0.5)
    # 左からstep_pxごとに色を見ていって、アクティブなタブを見つけたら、引っ張り出す
    current_x = start_pixel[0]
    while current_x < win_size[0]:
        pix = pyautogui.pixel(current_x, start_pixel[1])
        if pix == (103, 113, 135):
            pyautogui.moveTo(current_x, start_pixel[1])
            pyautogui.dragTo(*center_pix ,0.2)
            time.sleep(1)
            hwnd = win32gui.WindowFromPoint(center_pix)
            rect_pos = (sub_window_rect[0], sub_window_rect[1])
            rect_size = (sub_window_rect[2] - sub_window_rect[0], sub_window_rect[3] - sub_window_rect[1])

            win32gui.MoveWindow(hwnd, *rect_pos,*rect_size, True)
            break
        current_x += step_px
    keyboard.send(resize_area_hotkey)
    print('新規ウィンドウを作成しました')


if __name__ == '__main__':
    main()