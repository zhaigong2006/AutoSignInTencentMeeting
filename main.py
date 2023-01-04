# coding=<encoding name> ： # coding=utf-8
import os
import pyautogui
import time
from datetime import datetime
import cv2
import xlrd
import keyboard


def imgAutoClick(tempFile, whatDo, debug=False):  # 识别图标并且点击
    """
        temFile :需要匹配的小图
        whatDo  :需要的操作
                pyautogui.moveTo(w/2, h/2)# 基本移动
                pyautogui.click()  # 左键单击
                pyautogui.doubleClick()  # 左键双击
                pyautogui.rightClick() # 右键单击
                pyautogui.middleClick() # 中键单击
                pyautogui.tripleClick() # 鼠标当前位置3击
                pyautogui.scroll(10) # 滚轮往上滚10， 注意方向， 负值往下滑
        更多详情：https://blog.csdn.net/weixin_43430036/article/details/84650938
        debug   :是否开启显示调试窗口
    """
    pyautogui.screenshot('big.png')
    gray = cv2.imread("big.png", 0)
    img_template = cv2.imread(tempFile, 0)
    w, h = img_template.shape[::-1]
    res = cv2.matchTemplate(gray, img_template, cv2.TM_SQDIFF)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    top = min_loc[0]
    left = min_loc[1]
    x = [top, left, w, h]
    top_left = min_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)
    pyautogui.moveTo(top + h / 2, left + w / 2)
    whatDo(x)

    if debug:
        img = cv2.imread("big.png", 1)
        cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed", img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
    os.remove("big.png")


def signIn(meeting_id):  # 进入会议
    """
    本模块主要引入腾讯会议号，进入会议之中；
    """
    # 这一步是使用腾讯会议的绝对路径调用并启动腾讯会议
    os.startfile("C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe")  # 这是我腾讯会议的地址，如果你没有选择默认安装地址请更换成自己的
    time.sleep(7)  # 等待启动
    imgAutoClick("joinbtn.png", pyautogui.click, False)
    time.sleep(1)
    imgAutoClick("meeting_id.png", pyautogui.click, False)
    for i in range(0, 10):
        pyautogui.typewrite(["backspace"])
    time.sleep(1)
    pyautogui.write(meeting_id)
    time.sleep(2)
    imgAutoClick("final.png", pyautogui.click, False)
    time.sleep(1)


def signInAgain(meeting_id):  # 先退出会议再进入会议
    imgAutoClick("joinbtn.png", pyautogui.click, False)
    time.sleep(1)
    imgAutoClick("meeting_id.png", pyautogui.click, False)
    for i in range(0, 10):
        pyautogui.typewrite(["backspace"])
    time.sleep(1)
    pyautogui.write(meeting_id)
    time.sleep(2)
    imgAutoClick("final.png", pyautogui.click, False)
    time.sleep(1)


def signOut():  # 退出会议
    imgAutoClick("close.png", pyautogui.click, False)
    time.sleep(1)
    imgAutoClick("leave.png", pyautogui.click, False)
    time.sleep(1)


def excel(row, col):  # 从Excel表格中获得某行某列的数据（从零开始）
    excel_workbook = xlrd.open_workbook('table.xls')  # 打开table.xls，注意使用xlrd库只能使用xls格式的excel
    excel_worksheet = excel_workbook.sheet_by_index(0)
    number = excel_worksheet.cell_value(row, col)
    # print(excel_worksheet.cell_value(col, row))
    return number


def section():  # 获得当前时间段，返回对应时间段数字
    now = datetime.now().strftime("%m-%d-%H:%M")
    # now = "01-03-13:22"
    if int(now[6]) == 0 and int(now[7]) == 8 and int(now[9]) < 4 or (
            int(now[6]) == 0 and int(now[7]) == 7):
        return 1
    if (int(now[6]) == 0 and int(now[7]) == 8 and int(now[9]) >= 4) or (
            int(now[6]) == 0 and int(now[7]) == 9 and int(now[9]) < 3):
        return 2
    if (int(now[6]) == 0 and int(now[7]) == 9 and int(now[9]) >= 3) or (
            int(now[6]) == 1 and int(now[7]) == 0 and int(now[9]) < 4):
        return 3
    if (int(now[6]) == 1 and int(now[7]) == 0 and int(now[9]) >= 4) or (
            int(now[6]) == 1 and int(now[7]) == 1 and int(now[9]) < 3):
        return 4
    if (int(now[6]) == 1 and int(now[7]) == 1 and int(now[9]) >= 3) or (
            int(now[6]) == 1 and int(now[7]) == 3 and int(now[9]) < 4) or (int(now[6]) == 1 and int(now[7]) == 2):
        return 5
    if (int(now[6]) == 1 and int(now[7]) == 3 and int(now[9]) >= 4) or (
            int(now[6]) == 1 and int(now[7]) == 4 and int(now[9]) < 3):
        return 6


def now_number():  # 获得当前时段的会议码
    meeting_number = [0, 0, 0, 0, 0, 0]
    for i in range(0, 6):
        meeting_number[i] = int(excel(i, 7)).__str__()  # 获得Excel中的会议码
        # print(meeting_number)
    dayOfWeek = int(datetime.now().isoweekday())  # 获得今日星期几（如星期二返回2）
    print("今天是星期", dayOfWeek)
    final = meeting_number[int(excel(section(), dayOfWeek)) - 1]  # 根据今天星期几和现在的时间段获取现在应该输入的会议码
    print("正在进入：", final)
    return final


def load():  # 根据当前时间进入相应会议
    print("正在加载...")
    signOut()
    signIn(now_number())
    time.sleep(2)
    print('signed in!')


def manual_load(manual):  # 手动输入数字进入相应的会议
    print("正在加载：", manual)
    signOut()
    print(manual)
    final = str(int(excel(int(manual) - 1, 7)))
    print(final)
    signIn(final)
    time.sleep(2)
    print('signed in!')


print(
    "这个程序能够帮助您：自动根据当前时间切换到相应的腾讯会议室，若为下课时间视作下一节课的时间。按下ctrl+s就可以自动切换会议室。输入ctrl+1-6并按下回车可以手动切换到相应的会议室（123是语数英，456是D1D2D3）。如果你不是八班同学，请先根据自己的课程表修改table.xls，并确保table.xls与本程序在同一目录。如果程序图标在任务栏中闪烁但没有出现在屏幕上，请手动点一下程序图标。程序运行可能需要一定时间，请耐心等待！")
keyboard.add_hotkey("ctrl+s", lambda: load())
keyboard.add_hotkey("ctrl+1", lambda: manual_load(1))
keyboard.add_hotkey("ctrl+2", lambda: manual_load(2))
keyboard.add_hotkey("ctrl+3", lambda: manual_load(3))
keyboard.add_hotkey("ctrl+4", lambda: manual_load(4))
keyboard.add_hotkey("ctrl+5", lambda: manual_load(5))
keyboard.add_hotkey("ctrl+6", lambda: manual_load(6))

keyboard.wait()
