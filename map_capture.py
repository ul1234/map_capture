#!/usr/bin/python
# -*- coding: utf-8 -*-

import win32api
import win32gui
import win32con
import time
import Tkinter as tk
from PIL import Image, ImageGrab


class MapCap(tk.Toplevel):
    def __init__(self, master):
        tk.Toplevel.__init__(self, master)
        #self.overrideredirect(True)                 # 去掉顶层那一栏
        self.protocol('WM_DELETE_WINDOW', self.quit)
        self.geometry('900x600+400+200')

        text = 'How to use:\n\n1.change this window size by grabbing to set the default 1x map size.\n\n'
        text += '2.set width_n and heigth_n value.\n  The final map is then (width_n*height_n)\'s 1x size big.\n\n'
        text += '3.set time_wait value.\n  If the network is slow, increase this value.\n\n'
        text += '4.put this window on the top of the baidu map.\n\n'
        text += '5.press the button and just wait.\n\n'
        text += '6.please check "test.jpg" in the current folder when this window reappears.\n\n'
        text += 'Enjoy!!!'
        self.label = tk.Label(self, text = text, anchor = 'w', justify = 'left', font = ('courier', 12, 'bold'), fg = 'blue')
        self.label.pack(side = 'top', fill = 'both', padx = 10, pady = 10)

        self.width_var = tk.StringVar()
        self.height_var = tk.StringVar()
        self.time_wait_var = tk.StringVar()

        frame1 = tk.Frame(self)
        frame1.pack(side = 'top')

        tk.Label(frame1, text = 'width_n ', font = ('courier', 12, 'bold'), fg = 'red').pack(side = 'left', padx = 10, pady = 10)
        width_entry = tk.Entry(frame1, textvariable = self.width_var, font = ('courier', 12, 'bold'), fg = 'blue', width = 5)
        width_entry.pack(side = 'left', padx = 10, pady = 10)

        frame2 = tk.Frame(self)
        frame2.pack(side = 'top')

        tk.Label(frame2, text = 'height_n', font = ('courier', 12, 'bold'), fg = 'red').pack(side = 'left', padx = 10, pady = 10)
        height_entry = tk.Entry(frame2, textvariable = self.height_var, font = ('courier', 12, 'bold'), fg = 'blue', width = 5)
        height_entry.pack(side = 'left', padx = 10, pady = 10)

        frame3 = tk.Frame(self)
        frame3.pack(side = 'top')

        tk.Label(frame3, text = 'time_wait(s)', font = ('courier', 12, 'bold'), fg = 'red').pack(side = 'left', padx = 10, pady = 10)
        time_wait_entry = tk.Entry(frame3, textvariable = self.time_wait_var, font = ('courier', 12, 'bold'), fg = 'blue', width = 5)
        time_wait_entry.pack(side = 'left', padx = 10, pady = 10)

        self.button = tk.Button(self, text = 'Start Cap Map', font = ('courier', 12, 'bold'), fg = 'red', command = self.start_cap)
        self.button.pack(padx = 10, pady = 10)

        self.width_var.set('3')
        self.height_var.set('3')
        self.time_wait_var.set('2')

    def start_cap(self):
        self._get_window_screen_pos()
        width_n = int(self.width_var.get())
        height_n = int(self.height_var.get())
        time_wait = int(self.time_wait_var.get())
        self.composite_map(width_n, height_n, time_wait)

    def _get_window_screen_pos(self):
        self.cap_size = (self.winfo_width(), self.winfo_height())
        self.cap_pos = win32gui.ClientToScreen(self.winfo_id(), (0, 0)) \
                     + win32gui.ClientToScreen(self.winfo_id(), self.cap_size)
        print self.cap_pos

    def cap_screen(self):
        return ImageGrab.grab(self.cap_pos)

    def _move_map(self, map_win, dir):
        if dir == 'right':
            from_pos, to_pos = (self.cap_pos[2], self.cap_pos[1]), (self.cap_pos[0], self.cap_pos[1])
        elif dir == 'left':
            from_pos, to_pos = (self.cap_pos[0], self.cap_pos[1]), (self.cap_pos[2], self.cap_pos[1])
        elif dir == 'up':
            from_pos, to_pos = (self.cap_pos[0], self.cap_pos[1]), (self.cap_pos[0], self.cap_pos[3])
        else:
            from_pos, to_pos = (self.cap_pos[0], self.cap_pos[3]), (self.cap_pos[0], self.cap_pos[1])

        from_client_pos = win32gui.ScreenToClient(map_win, from_pos)
        to_client_pos = win32gui.ScreenToClient(map_win, to_pos)

        win32api.SetCursorPos(from_pos)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(0.5)

        win32api.SetCursorPos(to_pos)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        time.sleep(0.5)

    def composite_map(self, width_n, height_n, time_wait):
        #win32gui.ShowWindow(self.winfo_id(), win32con.SW_MINIMIZE)
        self.withdraw()
        time.sleep(1)
        map_win = win32gui.GetForegroundWindow()

        # move map to the left-top corner
        for i in range((width_n-1)/2):
            self._move_map(map_win, 'left')

        for i in range((height_n-1)/2):
            self._move_map(map_win, 'up')

        # cap map
        images = []
        h_direc = 'down'
        for w in range(width_n):
            for h in range(height_n):
                time.sleep(time_wait)
                image = self.cap_screen()
                real_h = h if h_direc == 'down' else height_n-h-1
                images.append((w, real_h, image))
                self._move_map(map_win, h_direc)
            self._move_map(map_win, 'right')
            h_direc = 'up' if h_direc == 'down' else 'down'
            self._move_map(map_win, h_direc)

        # restore map position
        for i in range((width_n-1)/2+1):
            self._move_map(map_win, 'left')

        for i in range((height_n-1)/2):
            self._move_map(map_win, h_direc)

        # omposite the map image
        composite_image = Image.new('RGB', (self.cap_size[0]*width_n, self.cap_size[1]*height_n), 'white')
        for w, h, image in images:
            composite_image.paste(image, (self.cap_size[0]*w, self.cap_size[1]*h))
        composite_image.save('test.jpg', quality = 200)

        self.deiconify()


if __name__ == '__main__':
    root = tk.Tk(className = ' Capture Map Tool')
    root.withdraw()
    cap = MapCap(root)
    cap.mainloop()

