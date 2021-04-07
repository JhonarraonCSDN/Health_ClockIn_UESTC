#!/usr/bin/python
# coding: utf-8
import os, sys, pyperclip, base64, cv2
import win32com.client
from time import sleep
import numpy as np


class ClockIn(object):
    def __init__(self, maxretry=100, sleep_time=0.1):
        info = ["人员类型：", "省份：", "城市：", "区县：", "详细地址：", "近14天健康情况：", "近14天有重点地区：", "近14天接触过重点地区：", "接触疑似：", "政府隔离：",
                "医学隔离：", "是否就诊：", "是否在校：", "家庭成员健康情况："]
        with open(os.getcwd() + "/Health_ClockIn.ini", encoding='utf-8') as f:
            info_text = f.readlines()
            self.renyuan = info_text[0].replace(info[0], "").replace('\n', '')
            self.shenfen = info_text[1].replace(info[1], "").replace('\n', '')
            self.chengshi = info_text[2].replace(info[2], "").replace('\n', '')
            self.quxian = info_text[3].replace(info[3], "").replace('\n', '')
            self.address = info_text[4].replace(info[4], "").replace('\n', '')
            self.jkqk = info_text[5].replace(info[5], "").replace('\n', '')
            self.zddq = info_text[6].replace(info[6], "").replace('\n', '')
            self.jczd = info_text[7].replace(info[7], "").replace('\n', '')
            self.jcys = info_text[8].replace(info[8], "").replace('\n', '')
            self.zfgl = info_text[9].replace(info[9], "").replace('\n', '')
            self.yxgl = info_text[10].replace(info[10], "").replace('\n', '')
            self.sfjz = info_text[11].replace(info[11], "").replace('\n', '')
            self.sfzx = info_text[12].replace(info[12], "").replace('\n', '')
            self.jtcy = info_text[13].replace(info[13], "").replace('\n', '')

        self.pic_path = os.getcwd() + "/pic"
        self.text_path = os.getcwd() + "/text" + "/text.txt"

        self.chrome_dir = info_text[14].replace("谷歌浏览器路径：", "").replace('\n', '').replace("\"","")
        print(self.chrome_dir)
        if not os.path.exists(self.chrome_dir):
            print("In ", __file__, " In", sys._getframe().f_lineno,
                  "无法找到Chorme浏览器的路径，请检查是否安装Chorme浏览器,并打开目录下Health_ClockIn.ini文件修改")
            input("按Enter键结束")
            exit()

        self.max_retry = maxretry
        self.sleep = sleep_time
        # 大漠插件
        self.dm = self.dm_init()
        # 参数修改
        print("填报信息如下：\n", "1、人员类型：", self.renyuan, '\n', "2、省份：", self.shenfen, '\n', "3、城市：", self.chengshi, '\n',
              "4、区县：", self.quxian, '\n', "5、详细地址：", self.address, '\n', "6、近14天健康情况：", self.jkqk, '\n',
              "7、近14天有重点地区:",
              self.zddq, '\n', "8、近14天接触过重点地区：", self.jczd, '\n', "9、接触疑似：", self.jcys, '\n', "10、政府隔离：", self.zfgl,
              '\n', "11、医学隔离：", self.yxgl, '\n', "12、是否就诊：", self.sfjz, '\n', "13、是否在校：", self.sfzx, '\n',
              "14、家庭成员健康情况：", self.jtcy, '\n', "请确认以上信息是否正确，如需修改请打开目录下Health_ClockIn.ini文件修改")
        sleep(5)
        # 屏幕分辨率
        self.screen_weight = self.dm.GetScreenWidth()
        self.screen_height = self.dm.GetScreenHeight()
        print("屏幕分辨率：", self.screen_weight, self.screen_height)

        if not os.path.exists(self.pic_path + '/' + str(self.screen_weight) + "_" + str(self.screen_height)):
            os.mkdir(self.pic_path + '/' + str(self.screen_weight) + "_" + str(self.screen_height))
        self.pic_path = self.pic_path + '/' + str(self.screen_weight) + "_" + str(self.screen_height)
        # 大漠插件加载图片
        self.dm.SetPath(self.pic_path)
        all_pic = ""
        for file in os.listdir(self.pic_path):
            if ".bmp" in file:
                all_pic += file
                all_pic += "|"
        self.dm.LoadPic(all_pic)
        self.dm.LoadPic(all_pic)
        self.hwnd = self.open_webdriver("统一身份认证")

    def dm_init(self):
        dm = 0
        try:
            dm = win32com.client.Dispatch('dm.dmsoft')
        except:
            dm = 0
        try:
            if dm == 0:
                os.system(r'regsvr32 /s %s\dm.dll' % os.getcwd())
                dm = win32com.client.Dispatch('dm.dmsoft')
        except:
            dm = 0
        try:
            if dm == 0:
                dm_path = os.getcwd() + "/注册大漠插件.bat"
                os.startfile(dm_path)
                dm = win32com.client.Dispatch('dm.dmsoft')
        except:
            input("无法注册大漠插件,按enter退出")
            exit()
        if dm != 0:
            return dm
        else:
            input("无法注册大漠插件,按enter退出")
            exit()

    def open_webdriver(self, text):
        def get_hwnd():
            hwnd = self.dm.EnumWindow(0, "", "Chrome_WidgetWin_1", 2).split(',')
            if hwnd[0] != '':
                for i in range(len(hwnd)):
                    title = self.dm.GetWindowTitle(hwnd[i])
                    if text in title:
                        self.dm.SetWindowState(hwnd[i], 1)
                        self.dm.SetWindowState(hwnd[i], 4)
                        return hwnd[i]
                for i in range(len(hwnd)):
                    self.dm.SetWindowState(hwnd[i], 1)
                    self.dm.SetWindowState(hwnd[i], 4)
                    if self.dm.GetWindowState(hwnd[i], 1):
                        return hwnd[i]
            else:
                os.startfile(self.chrome_dir)
            return 0

        for i in range(self.max_retry):
            h = get_hwnd()
            sleep(self.sleep)
            if h:
                return h

        print("In ", __file__, " In", sys._getframe().f_lineno, "无法找到hwnd")
        exit()

    def find_target(self, big_img, small_img, aim_value=80):
        img = cv2.cvtColor(big_img, cv2.COLOR_BGR2GRAY)

        high = small_img.shape[0]

        big_img_bw = np.copy(big_img)
        big_img_bw[big_img_bw < 254] = 0

        width = big_img_bw.shape[1]

        for j in range(width):
            value = np.sum(big_img_bw[:, j] != 0)
            if value > aim_value:
                target = j
                break

        img_end = np.copy(img)
        img_end[:, target] = 255

        return target / width

    def find_pic(self, pic_name, w1, h1, w2, h2, max_retry=10, sleep_time=0.1, acc=0.9, delta_color="000000",
                 is_center=1, is_fatal=0):
        for i in range(max_retry):
            target = self.dm.FindPic(w1, h1, w2, h2, pic_name, delta_color, acc, 0)
            if target[0] != -1:
                print("找到%s图片" % pic_name)
                x, y = target[1], target[2]
                if is_center:
                    center = self.dm.GetPicSize(pic_name).split(",")
                    w = int(center[0])
                    h = int(center[1])
                    return x + int(w / 2), y + int(h / 2)
                else:
                    return x, y
            else:
                sleep(sleep_time)

        if not is_fatal:
            print("未找到%s图片" % pic_name)
            return -1, -1
        else:
            print("未找到%s图片" % pic_name)
            exit()

    def find_pic_ex(self, pic_name, w1, h1, w2, h2, max_retry=10, sleep_time=0.1, acc=0.9, delta_color="9f2e3f-000000"):
        for i in range(max_retry):
            result = self.dm.FindPicEx(w1, h1, w2, h2, pic_name, delta_color, acc, 0)
            result = result.split("|")
            pic_names = pic_name.split("|")
            if result[0] != '':
                for i in range(len(result)):
                    tmp = result[i].split(",")
                    result[i] = tmp[1:]
                    center = self.dm.GetPicSize(pic_names[int(tmp[0])]).split(",")
                    w = int(center[0])
                    h = int(center[1])
                    result[i][0] = int(result[i][0]) + int(w / 2)
                    result[i][1] = int(result[i][1]) + int(h / 2)
                print("找到%d张%s图片" % (len(result), pic_name))
                return result
            else:
                sleep(sleep_time)
        if result[0] == '':
            print("未找到%s图片" % pic_name)
            result = []
            return result

    def click(self, x, y, sleep_time=0.1):
        self.dm.moveto(x, y)
        sleep(sleep_time)
        self.dm.LeftClick()
        sleep(sleep_time)
        self.dm.moveto(0, 0)  # 移开鼠标，以免影响识别
        sleep(sleep_time)

    def write_info(self, w1, h1, w2, h2, sleep_time=0.1):
        def paste(text):
            pyperclip.copy(text)
            self.dm.KeyDownChar("ctrl")
            self.dm.KeyDownChar("a")
            self.dm.KeyUpChar("a")
            self.dm.KeyDownChar("v")
            self.dm.KeyUpChar("v")
            self.dm.KeyUpChar("ctrl")
            self.dm.KeyDownChar("enter")
            self.dm.KeyUpChar("enter")

        def find_choice(target_name, choice_name, ):
            target_pic = target_name + ".bmp"
            x, y = self.find_pic(target_pic, w1, h1, w2, h2)
            if x != -1 and y != -1:
                self.click(x + 150, y)
                target_pic = "chazhao.bmp"
                x1, y1 = self.find_pic(target_pic, w1, h1, w2, h2)
                if x1 != -1 and y1 != -1:
                    self.click(x1, y1)
                    paste(choice_name)
                    shift = {"1920_1080": 30, "2560_1440": 30}
                    self.click(x1, y1 + shift[str(self.screen_weight) + "_" + str(self.screen_height)])
                    print(target_name, "已填写：", choice_name)
            else:
                print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)

        find_choice("renyuan", self.renyuan)
        find_choice("shenfen", self.shenfen)
        find_choice("chengshi", self.chengshi)
        find_choice("quxian", self.quxian)

        target_pic = "dizhi.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x + 150, y)
            paste(self.address)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)

        find_choice("jkqk", self.jkqk)
        find_choice("zddq", self.zddq)
        find_choice("jczd", self.jczd)
        find_choice("jcys", self.jcys)
        find_choice("zfgl", self.zfgl)
        find_choice("yxgl", self.yxgl)
        find_choice("sfjz", self.sfjz)
        find_choice("sfzx", self.sfzx)
        find_choice("jtcy", self.jtcy)

        target_pic = "bc.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        target_pic = "qr.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)

    def run_tasks(self, sleep_time=0.5):
        rect = self.dm.GetWindowRect(self.hwnd)
        w1, h1, w2, h2 = rect[1], rect[2], rect[3], rect[4]
        target_pic = "Lock.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x + 10, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        sleep(sleep_time)
        target_pic = "Login.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        sleep(sleep_time)
        target_pic = "Login1.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        sleep(sleep_time)
        target_pic = "Yanzhen.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2, is_center=0)  # 请完成安全验证-坐标点x，y
        if x != -1 and y != -1:
            def Yanzhen():
                self.dm.KeyPress(123)  # 按F12打开DevTools
                # 清除当前浏览器图片缓存
                target_pic = "clear.bmp"
                x1, y1 = self.find_pic(target_pic, w1, h1, w2, h2)
                self.click(x1, y1)
                # 点击重试按钮以获得验证图片
                target_pic = "pt.bmp"
                x1, y1 = self.find_pic(target_pic, w1, h1, w2, h2, is_fatal=1)  # 向右滑动填充拼图-坐标点x1，y1
                target_pic = "close.bmp"
                x2, y2 = self.find_pic(target_pic, w1, h1, w2, h2, is_fatal=1)  # 关闭图标-坐标点x2，y2
                target_x = int(x + (x2 - x) * 249 / 264)
                target_y = int(y + (y1 - y) * 67 / 223)
                self.click(target_x, target_y)
                self.dm.KeyPress(123)
                target_pic = "data.bmp"
                result = self.find_pic_ex(target_pic, w1, h1, w2, h2)
                if len(result) > 2:
                    print("In ", __file__, " In", sys._getframe().f_lineno, "找到%d张" % (len(result)), target_pic)
                    return 0
                elif len(result) == 2:
                    for i in range(len(result)):
                        x_r, y_r = result[i][0], result[i][1]
                        self.dm.moveto(x_r, y_r)
                        self.dm.RightClick()
                        target_pic = "copy.bmp"
                        x3, y3 = self.find_pic(target_pic, w1, h1, w2, h2)  # Copy-坐标点x3，y3
                        self.click(x3, y3)

                        target_pic = "copy_link.bmp"
                        x3, y3 = self.find_pic(target_pic, w1, h1, w2, h2)  # Copy_link_adress-坐标点x3，y3
                        self.click(x3, y3)

                        base64_str = pyperclip.paste().split(",")[-1]
                        imgString = base64.b64decode(base64_str)
                        nparr = np.fromstring(imgString, np.uint8)
                        image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                        cv2.imwrite(self.pic_path + "/img" + str(i) + ".png", image)
                    res = self.find_target(cv2.imread(self.pic_path + "/img0.png"),
                                           cv2.imread(self.pic_path + "/img1.png"), 30)

                    def drag():
                        x1 = x + 15
                        y1 = y + 220
                        target_x = x + int(res * 275) + 20
                        self.dm.moveto(x1, y1)
                        self.dm.LeftDown()
                        sleep(sleep_time)
                        self.dm.moveto(target_x, y1)
                        sleep(sleep_time)
                        self.dm.LeftUp()

                    drag()
                elif len(result) == 1:
                    print("In ", __file__, " In", sys._getframe().f_lineno, "只找到1张", target_pic)
                    return 0
                else:
                    print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
                    return 0
                target_pic = "JK.bmp"
                x1, y1 = self.find_pic(target_pic, w1, h1, w2, h2, max_retry=100, sleep_time=1)
                if x1 != -1 and y1 != -1:
                    return 1
                else:
                    return 0

            for i in range(self.max_retry):
                if Yanzhen():
                    print("验证成功")
                    break
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        sleep(sleep_time)
        target_pic = "Jk.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        sleep(sleep_time)
        target_pic = "xinzhen.bmp"
        x, y = self.find_pic(target_pic, w1, h1, w2, h2)
        if x != -1 and y != -1:
            self.click(x, y)
        else:
            print("In ", __file__, " In", sys._getframe().f_lineno, "未找到", target_pic)
        sleep(sleep_time)
        self.write_info(w1, h1, w2, h2)


if __name__ == "__main__":
    daily_tasks = ClockIn()
    daily_tasks.run_tasks()
