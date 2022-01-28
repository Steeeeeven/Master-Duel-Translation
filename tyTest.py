import sys
import tkinter
from  tkinter import messagebox
import win32api,win32gui,win32ui,win32con,win32print, keyboard
from pynput.mouse import Button, Controller, Listener
root = tkinter.Tk()

#获取真实的分辨率宽
real_wide = win32print.GetDeviceCaps(win32gui.GetDC(0), win32con.DESKTOPHORZRES)

#获取缩放后的分辨率宽
screen_wide = win32api.GetSystemMetrics(0) 
proportion = round(real_wide / screen_wide, 2)

#决斗界面窗口位置
window_duel_x = 29/1920
window_duel_y = 458/1080
window_duel_w = 371/1920
window_duel_h = 397/1080

#deck界面窗口位置
window_deck_x = 54/1920
window_deck_y = 464/1080
window_deck_w = 400/1920
window_deck_h = 333/1080

string = "400x333+50+300"
root.geometry(string)

    
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        listener.stop()
        root.destroy()        
        sys.exit(0)

#卡牌名称
name = tkinter.StringVar()
#空行(真的需要这样写吗)
blankLine = tkinter.StringVar()
#卡片内容
cardContain = tkinter.StringVar()

name.set("Def")
blankLine.set("")
cardContain.set("DefaultContain")

nameLabel = tkinter.Label(root, textvariable=name,font=("微软雅黑", 20, "bold")).pack(padx=10)
cardContainLabel = tkinter.Label(root, textvariable=cardContain,wraplength=80).pack()
 
def updatingTest():
    cardContain.set("DefauContainDefauContainDefauContainDefauContain")
    
def on_click( x, y, button, pressed):
    print ('Mouse clicked at ({0}, {1}) with {2}'.format(x, y, button))

def setLocation():
    md_process_window_name='masterduel'
    gamehwnd = win32gui.FindWindow(None, md_process_window_name)
    rect = win32gui.GetWindowRect(gamehwnd)
    print(rect[0],rect[1],rect[2],rect[3])
    left = rect[0]
    top = rect[1]
    gamew = rect[2] - left
    gameh = rect[3] - top
    
    print()
    x = (int)(rect[0]+(rect[2]-rect[0])*window_deck_x)
    y = (int)(rect[1]+(rect[3]-rect[1])*window_deck_y)
    w = (int)((rect[2]-rect[0])*window_deck_w)
    h = (int)((rect[3]-rect[1])*window_deck_h)

    string = f"{int(w)}x{int(h)}+{int(x)}+{int(y)}"
    print(string)
    root.geometry(string)
    root.update()

listener = Listener(on_click=on_click)
listener.start()
   
keyboard.add_hotkey('f1',updatingTest)
keyboard.add_hotkey('f2',setLocation)

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        listener.stop()
        root.destroy()        
        sys.exit(0)

        
exitButton = tkinter.Button(root, text ="关闭",anchor="sw", command = on_closing).place(x=10,y=10)


root.attributes("-topmost", True)
root.overrideredirect(True)
root.protocol('WM_DELETE_WINDOW', on_closing)
root.mainloop()