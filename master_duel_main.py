import os
import sys
import time
import tkinter
from  tkinter import messagebox
from PIL import Image,ImageFile
import dhash
import sqlite3
import pywintypes
import win32api,win32process,win32gui,win32ui,win32con,win32print
from ctypes import windll
import keyboard
from pynput.mouse import Button, Controller
from pynput.mouse import Listener

root = tkinter.Tk()

#获取真实的分辨率宽
real_wide = win32print.GetDeviceCaps(win32gui.GetDC(0), win32con.DESKTOPHORZRES)

#获取缩放后的分辨率宽
screen_wide = win32api.GetSystemMetrics(0) 
proportion = round(real_wide / screen_wide, 2)
print(f"缩放比为{proportion}")
#翻译窗口位置
window_x = 29/proportion
window_y = 458/proportion
window_w = 371/proportion
window_h = 397/proportion

string = f"{int(window_w)}x{int(window_h)}+{int(window_x)}+{int(window_y)}"
print(string)
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

#刷新
def updatingTest():
    nameLabel = tkinter.Label(root, textvariable=name,font=("微软雅黑", 20, "bold")).pack()
    blankLineLabel = tkinter.Label(root, textvariable=blankLine).pack()
    cardContainLabel = tkinter.Label(root, textvariable=cardContain,wraplength=window_w-20).pack()

def updatingTest1():
    window_x = 100
    string = f"{int(window_w)}x{int(window_h)}+{int(window_x)}+{int(window_y)}"
    root.geometry(string)
    root.update()
    cardContain.set("DefauContainDefauContainDefauContainDefauContain")
    
def on_click( x, y, button, pressed):
    print ('Mouse clicked at ({0}, {1}) with {2}'.format(x, y, button))









ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None
SetForegroundWindow = windll.user32.SetForegroundWindow

_img_p=['.png','.jpg']

#卡片图像位置y
y_1=(130/700)
y_2=(490/700)

#卡片图像位置x
x_1=(60/480)
x_2=(424/480)

#决定每次检测之后显示最相似的条目总数(避免可能的识别误差)
show_search_limit=1

md_process_name='masterduel.exe'
md_process_window_name='masterduel'
title = win32gui.GetWindowText (win32gui.GetForegroundWindow())
#for regenerate card image
fileDir='./origin_ygo_img'

c_dhash_dir='./card_image_check.db'
c_ygo_dir='./cards.cdb'

#screen_shot for 1920X1080
#shot where card image locate
#deck
deck_left_top=(64,200)
deck_right_bottom=(64+144,200+210)
#duel
duel_left_top=(40,208)
duel_right_bottom=(40+168,208+244)

def cls():
    os.system('cls' if os.name=='nt' else 'clear')
    
def hammingDist(s1, s2):
    assert len(s1) == len(s2)
    return sum([ch1 != ch2 for ch1, ch2 in zip(s1, s2)])
    
#thishwnd = win32gui.FindWindow(None, title)
#win32gui.SetWindowPos(thishwnd,win32con.HWND_NOTOPMOST,0, 0, 440, 600, 0)

def getFileList(dir, fileList):
    newDir = dir
    if os.path.isfile(dir):
        fileList.append(dir)
    elif os.path.isdir(dir):
        for s in os.listdir(dir):
            if os.path.splitext(s)[-1] not in _img_p:
                continue
            newDir = os.path.join(dir, s)
            getFileList(newDir, fileList)
    return fileList

def generate_card_img_basic_dhash(_list):
    conn = sqlite3.connect(c_dhash_dir)
    c = conn.cursor()
    
    c.execute(''' SELECT count(name) FROM sqlite_master WHERE type='table' AND name='CardDhash' ''')
    
    if c.fetchone()[0]!=1 : 
        conn.execute(
            '''
            CREATE TABLE CardDhash
            (id       INTEGER   PRIMARY KEY AUTOINCREMENT,
            code      TEXT  NOT NULL,
            dhash     TEXT  NOT NULL
            );'''
        )
    
    c.execute(''' SELECT count(*) FROM CardDhash ''')
    if c.fetchone()[0]==0:
        counter=0
        for _img_path in _list:
            _img=Image.open(_img_path)
            
            _y_1=int(_img.height*y_1)
            _y_2=int(_img.height*y_2)
            _x_1= int(_img.width*x_1)
            _x_2=int(_img.width*x_2)

            _img = _img.crop((_x_1,_y_1,_x_2,_y_2))
            row, col = dhash.dhash_row_col(_img)
            
            _img.close()
            
            _temp_dhash=dhash.format_hex(row, col)
                
            if _temp_dhash is None:
                print(f'Unbale read {_img_path},next')
                continue
            counter+=1
            _file_name=os.path.basename(_img_path).split('.')[0]
            
            # _cache.append({
            #     'code':_file_name,
            #     'dhash':_temp_dhash
            # })   
            conn.execute(f"INSERT INTO CardDhash (code,dhash) VALUES ('{_file_name}', '{_temp_dhash}' )");

            print(f"{counter} time,generate card {_file_name} dhash {_temp_dhash}")
        print("generate done")
        conn.commit()
    conn.close()

def get_card_img_dhash_cache():
    conn = sqlite3.connect(c_dhash_dir)
    c = conn.cursor()
    
    c.execute(''' SELECT count(name) FROM sqlite_master WHERE type='table' AND name='CardDhash' ''')
    
    if c.fetchone()[0]!=1 :  
        print("No table find")
        conn.close()
        return None
    c.execute(''' SELECT count(*) FROM CardDhash ''')
    if c.fetchone()[0]==0:
        print("No data Init")
        conn.close()
        return None
    
    cache=[]
    cursor = conn.execute("SELECT code,dhash from CardDhash")
    for row in cursor:
        cache.append(
            {
                'code':row[0],
                'dhash':row[1]
            }
        )

    conn.close()
    return cache

def get_game_window_info():
    hwnd=win32gui.FindWindow(0,md_process_window_name)
    return hwnd 

def window_shot_image(hwnd:int):
    app = win32gui.GetWindowText(hwnd)
    if not hwnd or hwnd<=0 or len(app)==0:
        return False,'无法找到游戏进程,不进行操作'
    isiconic=win32gui.IsIconic(hwnd)
    if isiconic:
        return False,'游戏处于最小化窗口状态,无法获取屏幕图像,不执行操作'
    
    left, top, right, bot = win32gui.GetClientRect(hwnd)
    w = right - left
    h = bot - top
    
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)
    
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    if result != 1:
        return False,"无法创建屏幕图像缓存"
    hDC = win32gui.GetDC(0)
    real_wide = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES) #获取真实的分辨率宽
    screen_wide = win32api.GetSystemMetrics(0) #获取缩放后的分辨率宽
    proportion = round(real_wide / screen_wide, 2)
    print(f"缩放比例:({proportion})")
    w=w*proportion
    h=h*proportion
    print(f"Win32返回的游戏窗口分辨率({w}x{h})，相关坐标将会进行缩放，缩放比率将为({w/1920},{h/1080})")
    
    return True,{
        "image":im,
        "current_window_zoom":(w/1920,h/1080),
    }
      
      

def get_image_db_cache():
    generate_card_img_basic_dhash(getFileList(fileDir,[]))
    _db_image_cache=get_card_img_dhash_cache()
    return _db_image_cache

def on_top():
    thishwnd = win32gui.FindWindow(None, title)
    rect = win32gui.GetWindowRect(thishwnd)
    x1 = rect[0]
    y1 = rect[1]
    w1 = rect[2] - x1
    h1 = rect[3] - y1
    win32gui.SetWindowPos(thishwnd,win32con.HWND_TOPMOST,x1, y1, w1, h1, 0)
    
def non_top():
    thishwnd = win32gui.FindWindow(None, title)
    rect = win32gui.GetWindowRect(thishwnd)
    x1 = rect[0]
    y1 = rect[1]
    w1 = rect[2] - x1
    h1 = rect[3] - y1
    win32gui.SetWindowPos(thishwnd,win32con.HWND_NOTOPMOST,x1, y1, w1, h1, 0)


def cv_card_info_at_deck_room(debug:bool=False):
    hwnd=get_game_window_info()
    status,result=window_shot_image(hwnd)
    if not status:
        print(result)
        return None

    zoom_w=result['current_window_zoom'][0]
    zoom_h=result['current_window_zoom'][1]
    
    _crop_area=(int(deck_left_top[0]*zoom_w),
                int(deck_left_top[1]*zoom_h),
                int(deck_right_bottom[0]*zoom_w),
                int(deck_right_bottom[1]*zoom_h))

    _img=result['image'].crop(_crop_area)
    
    if debug:
        print("debug:store first crop deck card locate(first_crop_deck)")
        _img.save("./first_crop_deck.png")
    
    _y_1=int(_img.height*y_1)
    _y_2=int(_img.height*y_2)
    _x_1= int(_img.width*x_1)
    _x_2=int(_img.width*x_2)
    
    _img = _img.crop((_x_1,_y_1,_x_2,_y_2))
    
    if debug:
        print("debug:store second crop deck card locate(second_crop_deck)")
        _img.save("./second_crop_deck.png")
        
    row, col = dhash.dhash_row_col(_img)
    
    target_img_dhash=dhash.format_hex(row, col)
    
    return target_img_dhash

def cv_card_info_at_duel_room(debug:bool=False):
    
    hwnd=get_game_window_info()
    status,result=window_shot_image(hwnd)
    if not status:
        print(result)
        return None

    zoom_w=result['current_window_zoom'][0]
    zoom_h=result['current_window_zoom'][1]
    
    _crop_area=(int(duel_left_top[0]*zoom_w),
                int(duel_left_top[1]*zoom_h),
                int(duel_right_bottom[0]*zoom_w),
                int(duel_right_bottom[1]*zoom_h))

    _img=result['image'].crop(_crop_area)
    
    if debug:
        print("debug:store first crop duel card locate(first_crop_duel)")
        _img.save("./first_crop_duel.png")
    
    _y_1=int(_img.height*y_1)
    _y_2=int(_img.height*y_2)
    _x_1= int(_img.width*x_1)
    _x_2=int(_img.width*x_2)
    
    _img = _img.crop((_x_1,_y_1,_x_2,_y_2))
    
    if debug:
        print("debug:store second crop duel card locate(second_crop_duel)")
        _img.save("./second_crop_duel.png")
        
    row, col = dhash.dhash_row_col(_img)
    
    target_img_dhash=dhash.format_hex(row, col)
    
    return target_img_dhash

def translate(type:int,cache:list,debug:bool=False):
    get_game_window_info()
    if cache is None or len(cache)==0:
        print("无法读取图像指纹信息,不执行操作(card_image_check.db是不是12K，是的话这是个空库没数据的)")
        return
    cls()
    start_time=time.time()
    if type==1:
        print("翻译卡组卡片")
        dhash_info=cv_card_info_at_deck_room(debug)
    elif type==2:
        print("翻译决斗卡片")
        dhash_info=cv_card_info_at_duel_room(debug)
    else:
        print("not support")
        return
    if not dhash_info:
        return
    
    results=[]
    
    for _img_dhash in cache:
        d_score = 1 - hammingDist(dhash_info,_img_dhash['dhash']) * 1. / (32 * 32 / 4)
        
        results.append({
                'card':_img_dhash['code'],
                'score':d_score
            })
                
        results.sort(key=lambda x:x['score'],reverse=True)
        
        if len(results)>show_search_limit:
            results=results[:show_search_limit]  
    end_time=time.time()
    
    ygo_sql=sqlite3.connect(c_ygo_dir)
    for card in results:
        try:
            cursor = ygo_sql.execute(f"SELECT name,desc from texts WHERE id='{card['card']}' LIMIT 1")
        except:
            print("读取ygo数据库异常,是不是没有将card.cdb放进来")
            return
        if cursor.arraysize!=1:
            print(f"card {card['card']} not found")
            ygo_sql.close()
            return
        data=cursor.fetchone()
        card['name']=data[0]
        card['desc']=data[1]
    ygo_sql.close()
    print('匹配用时: %.6f 秒'%(end_time-start_time))
    print(f"识别结果【匹配概率由高到低排序】")
    for card in results:
        if card['score']<0.93:
            print("警告:相似度匹配过低,可能游戏卡图与缓存库的版本卡图不同或未知原因截图区域错误\n修改enable_debug查看截取图片信息分析原因\n")
        print(f"{card['name']}(密码:{card['card']},相似度:{card['score']})\n{card['desc']}\n")
    print("-----------------------------------")
    print("shift+g翻译卡组卡片,shift+f翻译决斗中卡片,ctrl+q关闭\n请确保您已经点开了目标卡片的详细信息!!!")

def setLocation():
    gamehwnd = win32gui.FindWindow(None, md_process_window_name)
    rect = win32gui.GetWindowRect(gamehwnd)
    left = rect[0]
    top = rect[1]
    gamew = rect[2] - left
    gameh = rect[3] - top
    
    #这几个也要乘以缩放比
    print(left)
    print(top)
    zoom_w=gamew/1920
    zoom_h=gameh/1080
    x = (int)(left+zoom_w*window_x)
    y = top+zoom_h*window_y
    w = zoom_w*window_w
    h = zoom_h*window_h
    print(x)
    print(y)
    print(w)
    print(h)
    string = f"{int(w)}x{int(h)}+{int(x)}+{int(y)}"
    print(string)
    root.geometry(string)
    root.update()

    



cache=get_image_db_cache()
enable_debug=False
print("shift+g翻译卡组卡片,shift+f翻译决斗中卡片,ctrl+q关闭\n请确保您已经点开了目标卡片的详细信息!!!")
keyboard.add_hotkey('shift+g',translate,args=(1,cache,enable_debug))
keyboard.add_hotkey('shift+f',translate,args=(2,cache,enable_debug))
keyboard.add_hotkey('f1',on_top)
keyboard.add_hotkey('f2',non_top)
keyboard.add_hotkey('f3',setLocation)

listener = Listener(on_click=on_click)
listener.start()

root.attributes("-topmost", True)
root.protocol('WM_DELETE_WINDOW', on_closing)
root.mainloop()

