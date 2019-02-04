import random
import win32api,win32gui
import sys
import time
import win32con
from PIL import ImageGrab

#获取游戏窗口
class_name = "TMain"
title_name = "Minesweeper Arbiter "
handler = win32gui.FindWindow(class_name, title_name)

#窗口坐标
left = 0
up = 0
right = 0
down = 0

#窗口捕获提示及坐标打印
if handler:
    print('窗口捕获成功')
    left,up,right,down = win32gui.GetWindowRect(handler)
    print('窗口坐标')
    print(str(left)+' '+str(right)+' '+str(up)+' '+str(down))
else:
    print('未找到窗口')

#计算雷区范围
left += 15
up += 101
right -= 15
down -= 15

#抓取雷区图像
rect = (left,up,right,down)
img = ImageGrab.grab(rect)

#方块大小
n_width,n_height = 16,16
#方块数量
n_num_x = int((right - left) / n_width)
n_num_y = int((down - up) / n_height)

#方块的所有状态包含的RGB列表
n_untap = [(225, (192, 192, 192)), (31, (128, 128, 128))]   #空白，未被点击
n_boom_white = [(4, (255, 255, 255)), (144, (192, 192, 192)), (31, (128, 128, 128)), (77, (0, 0, 0))]  #被点击，有雷，白底
n_boom_red = [(4, (255, 255, 255)), (144, (255, 0, 0)), (31, (128, 128, 128)), (77, (0, 0, 0))]  #被点击，有雷，红底
n_flag = [(54, (255, 255, 255)), (17, (255, 0, 0)), (109, (192, 192, 192)), (54, (128, 128, 128)), (22, (0, 0, 0))] #插旗，有雷
num_0 = [(54, (255, 255, 255)), (148, (192, 192, 192)), (54, (128, 128, 128))]   #被点击，无雷，无数字，即数字0
num_1 = [(185, (192, 192, 192)), (31, (128, 128, 128)), (40, (0, 0, 255))]  #数字1
num_2 = [(160, (192, 192, 192)), (31, (128, 128, 128)), (65, (0, 128, 0))]  #数字2
num_3 = [(62, (255, 0, 0)), (163, (192, 192, 192)), (31, (128, 128, 128))]  #数字3
num_4 = [(169, (192, 192, 192)), (31, (128, 128, 128)), (56, (0, 0, 128))]  #数字4
num_5 = [(70, (128, 0, 0)), (155, (192, 192, 192)), (31, (128, 128, 128))]  #数字5
num_6 = [(153, (192, 192, 192)), (31, (128, 128, 128)), (72, (0, 128, 128))]    #数字6

#建立雷区状态的列表
map = [[0 for i in range(n_num_x)] for i in range(n_num_y)]
#游戏结束标志
gameover = 1

#扫描整个雷区，标记每个块的状态
def scanmap():
    global img
    global map
    for y in range(n_num_y):
        for x in range(n_num_x):
            now_img = img.crop((x * n_num_x, y * n_num_y, (x + 1) * n_num_x, (y + 1) * n_num_y))
            if now_img.getcolors() == num_0:
                map[y][x] = 0
            elif now_img.getcolors() == num_1:
                map[y][x] = 1
            elif now_img.getcolors() == num_2:
                map[y][x] = 2
            elif now_img.getcolors() == num_3:
                map[y][x] = 3
            elif now_img.getcolors() == num_4:
                map[y][x] = 4
            elif now_img.getcolors() == num_5:
                map[y][x] = 5
            elif now_img.getcolors() == num_6:
                map[y][x] = 6
            elif now_img.getcolors() == n_untap:
                map[y][x] = -1
            elif now_img.getcolors() == n_flag:
                map[y][x] = -2
            elif now_img.getcolors() == n_boom_red or now_img.getcolors() == n_boom_white:
                map[y][x] = -3
            
            else:
                print("无法识别图像")
                print("坐标")
                print((y,x))
                print("颜色")
                print(now_img.getcolors())
                sys.exit(0)

def calculate():
    #扫描每个方块
    for y in range(n_num_y):
        for x in range(n_num_x):
            #方块内为数字，统计周围各状态方块的数量
            if map[y][x] >= 1 and map[y][x] <= 5:
                untap_num = 0
                empty_num = 0
                flag_num = 0
                num_num = 0
                #周围方块各状态的数量
                for yy in range(y-1,y+2):
                    for xx in range(x-1,x+2):
                        #边界范围内才操作
                        if yy >= 0 and xx >= 0 and yy < n_num_y and xx < n_num_x:
                            #排除当前方块
                            if not (y == yy and x == xx):
                                if map[yy][xx] == 0:
                                    empty_num += 1
                                elif map[yy][xx] == -1:
                                    untap_num += 1
                                elif map[yy][xx] == -2:
                                    flag_num += 1
                                else:
                                    num_num += 1
            #对当前方块周围进行点击或插旗操作
            for yy in range(y-1,y+2):
                for xx in range(x-1,x+2):
                    #边界范围内才操作
                    if yy >= 0 and xx >= 0 and yy < n_num_y and xx < n_num_x:
                        #排除当前方块
                        if not (y == yy and x == xx):
                            #如果是未被点击的方块
                            if map[yy][xx] == -1:
                                #未被点击的方块等于炸弹数
                                if map[y][x] == empty_num:
                                    win32api.SetCursorPos([left + xx * n_width, up + yy * n_height])
                                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                                elif map[y][x] == flag_num or map[yy][xx] == empty_num + flag_num:
                                    win32api.SetCursorPos([left + xx * n_width, up + yy * n_height])
                                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            untap_num = 0
            empty_num = 0
            flag_num = 0
            num_num = 0

def main():
    win32api.SetCursorPos([left + 5 * n_width, up + 5 * n_height])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    while(1):
        scanmap()
        calculate()

if __name__ == "__main__":
    main()
    



                    


                                
                                    


            