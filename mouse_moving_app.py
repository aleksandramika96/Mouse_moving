# -*- coding: utf-8 -*-
"""
@author: mikson
"""

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
#from time import strftime
import win32api, win32con


MINUTES = list(range(1, 15+1))
APP_TITLE = 'Mouse moving app'
TEXT_FONT = 'Calibri'
FONT_SIZE = 11

SCREEN_WIDTH = win32api.GetSystemMetrics(0)
SCREEN_HEIGHT = win32api.GetSystemMetrics(1)
X = int(SCREEN_WIDTH/2) 
Y = int(SCREEN_HEIGHT/2)
LIST_WIDTH_HEIGHT = [(X, Y), (X, Y+100), (X+100, Y+100), (X+100, Y)] #position mouse moving 
DICT_WIDTH_HEIGHT = dict((i, j) for i, j in enumerate(LIST_WIDTH_HEIGHT))

class CountDownMessageBox():
    
    TEXT = 'Counting down'
    
    def __init__(self, app, messageText=TEXT):
        self.newWindow = tk.Toplevel(app.mainWindow, cursor='pirate') #additional window 
        self.messageText = messageText      
        self.period = int(app.period.get())*60
        self.build()
        
    def build(self):
        tk.Label(self.newWindow, bitmap='hourglass', padx=10, font=(TEXT_FONT,FONT_SIZE), compound='left',
                 text=self.messageText, wraplength=200, fg='black').grid(row=0, column=0)

        self.timer_var = tk.StringVar()       
        tk.Label(self.newWindow, textvariable=self.timer_var, font=(TEXT_FONT,FONT_SIZE), fg='red').grid(row=1, column=0, padx=20, pady=20)
        
        self.iterator = tk.IntVar()
        tk.Label(self.newWindow, textvariable=self.iterator, font=(TEXT_FONT,FONT_SIZE), fg='black').grid(row=2, column=0, padx=20, pady=20)
        
        self.count_down(self.period)
  
    def count_down(self, time_count , iterator=0):
        mins = time_count // 60
        secs = time_count % 60
        self.timer_var.set('Time {:02d}:{:02d}'.format(mins, secs))
        self.iterator.set('{} moves'.format(iterator))
        
        if time_count==0:
            self.move_cursor(iterator)
            time_count=self.period+1
            iterator+=1

        time_count-=1
        self.newWindow.after(1000, self.count_down, time_count, iterator)

    
    def move_cursor(self, iterator):
        w = DICT_WIDTH_HEIGHT.get(iterator%len(DICT_WIDTH_HEIGHT))[0]
        h = DICT_WIDTH_HEIGHT.get(iterator%len(DICT_WIDTH_HEIGHT))[1]
        win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE|win32con.MOUSEEVENTF_MOVE,
                             int(w/SCREEN_WIDTH * 65535.0),
                             int(h/SCREEN_HEIGHT * 65535.0))        

class App:
 
    def __init__(self, mainWindow):
        self.mainWindow = mainWindow    
        self.build()
         
    def build(self):
        self.mainFrame = tk.Frame(self.mainWindow)
        self.mainFrame.pack(fill='both', expand=True)
 
        StartApp = tk.Button(self.mainWindow, text='Start', command=self.on_button, height=1, width=22)
        StartApp.pack(expand=True, padx=40, pady=8)
        
        tk.Label(self.mainWindow, font=(TEXT_FONT,FONT_SIZE), text='Set minutes interval').pack()

        self.period = ttk.Combobox(self.mainWindow, height=1, width=20, values=MINUTES)
        self.period.current(0)
        self.period.pack()

        ExitApp = tk.Button(self.mainWindow, text='Stop', command=self.exit_app, height=1, width=22)
        ExitApp.pack(expand=True, padx=40, pady=8)
    
    def on_button(self):
        try:
            if int(self.period.get()) not in MINUTES:            
                raise ValueError              
            else: 
                self.newWindow = CountDownMessageBox(self)
        except ValueError:
            tk.messagebox.showinfo(message='Please, select a value from the list')
        
    def exit_app(self):
        MsgBox = tk.messagebox.askquestion('Exit Application','Are you sure you want to exit the application?', icon='warning') #, type = 'yesno'
        
        if MsgBox=='yes':
           self.mainWindow.destroy()
                    
def main():
    mainWindow = tk.Tk()
    mainWindow.title(APP_TITLE)
 
    app = App(mainWindow)     
    mainWindow.mainloop()
  
if __name__ == '__main__':
    main()
