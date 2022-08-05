from tabnanny import check
import tkinter
import time
import json
import os
import random
import win32api
import win32con
import win32gui_struct
import win32gui
import codecs
import threading
from PIL import ImageTk

KBjsonname = './配置文件/课表配置.json'
CSjsonname = './配置文件/参数配置.json'
GNjsonname = './配置文件/功能配置.json'
BJjsonname = './配置文件/班级配置.json'
with open(CSjsonname, encoding='utf-8') as f:
    listt = json.load(f)  # 获取json文件
    f.close()
font = font1 = (listt['font'], listt['lessonlarge'])
with open(BJjsonname, encoding='utf-8') as m:
    BJlist = json.load(m)  # 获取json文件
    f.close()
with open(KBjsonname, encoding='utf-8') as s:
    KBlist = json.load(s)  # 获取json文件
    f.close()
with open(GNjsonname,encoding='utf-8') as d:
    GNlist = json.load(d)
    d.close()
class SysTrayIcon(object):
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]

    FIRST_ID = 1023

    def __init__(self,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit=None,
                 default_menu_index=None,
                 window_class_name=None, ):

        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit

        menu_options = menu_options + (('退出', None, self.QUIT),)
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id

        self.default_menu_index = (default_menu_index or 0)
        self.window_class_name = window_class_name or "SysTrayIconPy"

        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
                       win32con.WM_DESTROY: self.destroy,
                       win32con.WM_COMMAND: self.command,
                       win32con.WM_USER + 20: self.notify, }
        # Register the Window class.
        window_class = win32gui.WNDCLASS()
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map  # could also specify a wndproc.
        classAtom = win32gui.RegisterClass(window_class)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(classAtom,
                                          self.window_class_name,
                                          style,
                                          0,
                                          0,
                                          win32con.CW_USEDEFAULT,
                                          win32con.CW_USEDEFAULT,
                                          0,
                                          0,
                                          hinst,
                                          None)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()

        win32gui.PumpMessages()

    def _add_ids_to_menu_options(self, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            else:
                print('Unknown item', option_text, option_icon, option_action)
            self._next_action_id += 1
        return result

    def refresh_icon(self):
        # Try and find a custom icon
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(self.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       self.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:
            print("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0,
                          win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                          win32con.WM_USER + 20,
                          hicon,
                          self.hover_text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)

    def restart(self, hwnd, msg, wparam, lparam):
        self.refresh_icon()

    def destroy(self, hwnd, msg, wparam, lparam):
        if self.on_quit: self.on_quit(self)
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
        elif lparam == win32con.WM_RBUTTONUP:
            self.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:
            pass
        return True

    def show_menu(self):
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        # win32gui.SetMenuDefaultItem(menu, 1000, 0)

        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                self.hwnd,
                                None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def create_menu(self, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = self.prep_menu_icon(option_icon)

            if option_id in self.menu_actions_by_id:
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(self, icon):
        # First load the icon.
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)

        return hbm

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)

    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)


class xiugai_kebiao:
    def init(i):
        global v1,v2,v3,v4,v5,v6,v7,roo
        roo = tkinter.Toplevel()
        roo.title('课表设置')
        v1 = tkinter.StringVar()
        v2 = tkinter.StringVar()
        v3 = tkinter.StringVar()
        v4 = tkinter.StringVar()
        v5 = tkinter.StringVar()
        v6 = tkinter.StringVar()
        v7 = tkinter.StringVar()

        t1 = tkinter.Label(roo,text='星期一').grid(row=1,column=1)
        t2 = tkinter.Label(roo,text='星期二').grid(row=2,column=1)
        t3 = tkinter.Label(roo,text='星期三').grid(row=3,column=1)
        t4 = tkinter.Label(roo,text='星期四').grid(row=4,column=1)
        t5 = tkinter.Label(roo,text='星期五').grid(row=5,column=1)
        t6 = tkinter.Label(roo,text='星期六').grid(row=6,column=1)
        t7 = tkinter.Label(roo,text='星期日').grid(row=7,column=1)
        E1 = tkinter.Entry(roo,textvariable=v1).grid(row=1,column=2)
        E2 = tkinter.Entry(roo,textvariable=v2).grid(row=2,column=2)
        E3 = tkinter.Entry(roo,textvariable=v3).grid(row=3,column=2)
        E4 = tkinter.Entry(roo,textvariable=v4).grid(row=4,column=2)
        E5 = tkinter.Entry(roo,textvariable=v5).grid(row=5,column=2)
        E6 = tkinter.Entry(roo,textvariable=v6).grid(row=6,column=2)
        E7 = tkinter.Entry(roo,textvariable=v7).grid(row=7,column=2)
        
        v1.set(KBlist['Monday'])
        v2.set(KBlist['Tuesday'])
        v3.set(KBlist['Wednesday'])
        v4.set(KBlist['Thursday'])
        v5.set(KBlist['Friday'])
        v6.set(KBlist['Saturday'])
        v7.set(KBlist['Sunday'])

        b1 = tkinter.Button(roo,text='修改',command=xiugai_kebiao.xiugai)
        b1.grid(row=8,column=1)
        roo.mainloop()
    def xiugai():
        print(1)
        with codecs.open(KBjsonname,'w','utf-8') as s:
            KBlist['Monday'] = v1.get()
            KBlist['Tuesday'] = v2.get()
            KBlist['Wednesday'] = v3.get()
            KBlist['Thursday'] = v4.get()
            KBlist['Friday'] = v5.get()
            KBlist['Saturday'] = v6.get()
            KBlist['Sunday'] = v7.get()
            json.dump(KBlist, s, ensure_ascii=False)
        s.close()
     

class xiugai_peizhi:
    global f, params
    with codecs.open(CSjsonname, 'rb', 'utf-8') as f:
        params = json.load(f)
    f.close

    def get_json_data(self):
        params[1][t.get()] = t1.get()
        print(params[1][t.get()])

    def write_json_data(self):
        with codecs.open('./配置文件/参数配置.json', 'w', 'utf-8') as r:
            json.dump(params, r, ensure_ascii=False)
        r.close()
        print('文件修改完成')

    def init(i):
        global t, t1
        roo1 = tkinter.Toplevel()
        roo1.title('设置')
        t = tkinter.StringVar()
        om = 1
        roo1.mainloop()


def restart(i):
    import subprocess
    subprocess.Popen(r'restart.exe')


class get_birth:
    
    def aaa(self):
        mouth = time.strftime('%m')
        timestr = time.strftime('%m%d')
        try:
            global str11
            name = BJlist[0][timestr]
            print(1)
            str11 = '根据大数据分析，今天是{}的生日。   '.format(name)
        except:
            windo.destroy()

    def get_birt():
        global qwe, windo,str11
        str11 = str11[1:] + str11[0]
        qwe.configure(text=str11)
        windo.after(100, get_birth.get_birt)

    def getbirth(self):
            global qwe, windo
            windo = tkinter.Tk()
            
            windo.title('生日提醒服务')
            windo.geometry("350x50-0-50")
            
            qwe = tkinter.Label(windo,
                                text='',
                                fg='black',
                                font=(listt['font'], 20))
            qwe.pack()
            get_birth.aaa(self)
            try:
                get_birth.get_birt()
            except:
                pass
            windo.mainloop()





class dian():
    global dataMat, is_run, going
    dataMat = BJlist[1]
    going = True
    is_run = False

    def lottery_roll(var1, var2):
        global going
        show_member = random.choice(dataMat)
        var1.set(show_member)
        if going:
            window.after(50, dian.lottery_roll, var1, var2)
        else:
            var2.set('恭喜 {} !!!'.format(show_member))
            going = True
            return

    def lottery_start(var1, var2):
        global is_run
        if is_run:
            return
        is_run = True
        var2.set('幸运儿是你吗。。。')
        dian.lottery_roll(var1, var2)

    def lottery_end():
        global going, is_run
        if is_run:
            going = False
            is_run = False


def dianming(self):
    global window
    window = tkinter.Toplevel()
    window.geometry('405x320+250+15')
    window.title('滚动点名器')
    bg_label = tkinter.Label(window, width=70, height=24, bg='#ECf5FF')
    bg_label.place(anchor=tkinter.NW, x=0, y=0)
    var1 = tkinter.StringVar(value='即 将 开 始')
    show_label1 = tkinter.Label(window, textvariable=var1, justify='left', anchor=tkinter.CENTER, width=17, height=3,
                                bg='#BFEFFF',
                                font='楷体 -40 bold', foreground='black')
    show_label1.place(anchor=tkinter.NW, x=21, y=20)
    var2 = tkinter.StringVar(value='幸运儿是你吗。。。')
    show_label2 = tkinter.Label(window, textvariable=var2, justify='left', anchor=tkinter.CENTER, width=38, height=3,
                                bg='#ECf5FF',
                                font='楷体 -18 bold', foreground='red')
    show_label2.place(anchor=tkinter.NW, x=21, y=240)
    button1 = tkinter.Button(window, text='开始', command=lambda: dian.lottery_start(var1, var2), width=14, height=2,
                             bg='#A8A8A8',
                             font='宋体 -18 bold')
    button1.place(anchor=tkinter.NW, x=20, y=175)
    button2 = tkinter.Button(window, text='结束', command=lambda: dian.lottery_end(), width=14, height=2, bg='#A8A8A8',
                             font='宋体 -18 bold')
    button2.place(anchor=tkinter.NW, x=232, y=175)
    window.mainloop()


def gettime():
    timestr = time.strftime("%H:%M:%S")
    lb.configure(text=timestr)
    root.after(1000, gettime)


def getlesson():
    week = time.strftime("%A")  # 获取今天是周几
    apm = time.strftime("%p")
    if week == "Sunday":
        lesson = listt[week]
    else:
        if apm == 'PM':
            lesson = KBlist[week][4:]  # 获取今日课表
        else:
            lesson = listt[week][:4]
    lb1.configure(text=lesson)
    root.after(1000, getlesson)


def ppt_close(i):
    os.system('taskkill /f /im wps.exe')


def jianpan_close(i):
    os.system('taskkill /f /im TextInputHost.exe')


def time_class():
    global root11, lb1, lb
    root11 = tkinter.Tk()
    # root.overrideredirect(True)
    root11.attributes('-alpha', listt[1]["alpha"])
    root11.attributes('-topmost', 1)
    root11.title('li')
    lb = tkinter.Label(root11, text='',
                       fg=listt[1]['fontcolor'],
                       font=(listt[1]['font'], listt[1]['timelarge']),
                       )
    lb1 = tkinter.Label(root11, text='',
                        fg=listt[1]['fontcolor'],
                        font=font1)
    bt1 = tkinter.Button(root11, text="ppt点不开点这里",
                         fg=listt[1]['fontcolor'],
                         font=font1,
                         command=ppt_close)
    bt2 = tkinter.Button(root11, text='点名',
                         fg=listt[1]['fontcolor'],
                         font=font1,
                         command=dianming)
    bt3 = tkinter.Button(root11, text='键盘打不开',
                         fg=listt[1]['fontcolor'],
                         font=font1,
                         command=jianpan_close)
    lb.pack()
    lb1.pack()
    bt1.pack()
    bt2.pack()
    bt3.pack()
    gettime()
    getlesson()
    root.mainloop()
class GN_XZ:
    def init(id):
        rt = tkinter.Toplevel()
        p = tkinter.PhotoImage(file="./图片/按钮关.png",height=20,width=40)
        l1 = tkinter.Label(rt,text='生日消息推送').grid(row=1,column=1)
        l2 = tkinter.Label(rt,text='开机自启动').grid(row=2,column=1)
        del_button = tkinter.Button(rt, text='DEL', width=20, height=20)
        del_icon = ImageTk.PhotoImage(resize(os.getcwd()+'/delete.png', 0))
        del_button.config(image=del_icon)
        del_button.bind('<Button-1>', delete_selected_image)
        del_button.grid(row=0, column=0, sticky=tkinter.W)
    def check():
        print(1)

if __name__ == '__main__':
    def tuop():
        icons = '123.ico'
        hover_text = "桌面控制中心"
        menu_options = (('设置课表', icons, xiugai_kebiao.init),
                        ('设置参数', icons, xiugai_peizhi.init),
                        ('重新启动', icons, restart),
                        ('ppt打不开',icons,ppt_close),
                        ('键盘打不开',icons,jianpan_close),
                        ('点名',icons,dianming),
                        ('功能设置',icons,GN_XZ.init)
                        )

        def bye(SysTrayIcon):
            os._exit(1)

        SysTrayIcon(icons, hover_text, menu_options, on_quit=bye)


    def bir():
        get_birth.getbirth(1)


    global lb1, lb
    global root
    threading.Thread(target=tuop).start()
    root = tkinter.Tk()
    if GNlist['birthdayTX'] == 'Ture':
        threading.Thread(target=bir).start()
    root.geometry("+0+0")
    root.attributes('-alpha', listt["alpha"])
    root.attributes('-topmost', 1)
    root.title('li')
    lb = tkinter.Label(root, text='',
                       fg=listt['fontcolor'],
                       font=(listt['font'], listt['timelarge']),
                       )
    lb1 = tkinter.Label(root, text='',
                        fg=listt['fontcolor'],
                        font=font1)
    lb.pack()
    lb1.pack()
    gettime()
    getlesson()
    root.mainloop()
