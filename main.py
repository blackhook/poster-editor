# -*- coding:utf-8 -*-
import importlib, sys, os, datetime, shutil, time
import configparser
import xlrd
import chardet

importlib.reload(sys)
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

logo = """
                 _                          _ _ _             
 _ __   ___  ___| |_ ___ _ __       ___  __| (_) |_ ___  _ __ 
| '_ \ / _ \/ __| __/ _ \ '__|____ / _ \/ _` | | __/ _ \| '__|
| |_) | (_) \__ \ ||  __/ | |_____|  __/ (_| | | || (_) | |   
| .__/ \___/|___/\__\___|_|        \___|\__,_|_|\__\___/|_|   
|_|                                                           

"""
print(logo)

iniconfig = configparser.ConfigParser()
fd = open('./config.ini', "rb")
buf = fd.read()
result = chardet.detect(buf)
iniconfig.read('./config.ini', encoding=result["encoding"])
numCount = 0
excel_path = iniconfig.get("source", "excel_path")
data = xlrd.open_workbook(excel_path)
data.sheet_names()
table = data.sheet_by_index(0)
names = table.col_values(0)[1:]
inputfont = iniconfig.get("source", "inputfont")
fontsize = int(iniconfig.get("image", "fontsize"))
font = ImageFont.truetype(inputfont, fontsize)
color = iniconfig.get("image", "color")
img_path = iniconfig.get("source", "img_path")
header_y = iniconfig.get("image", "header_y")
x_offset = int(iniconfig.get("image", "x_offset"))
out_path = iniconfig.get("image", "out_path")
img_type = iniconfig.get("image", "img_type")

print("宝贝，确认是要生成 " + str(names[0:5])[2:-2].replace("', '", ",") + " 一共" + str(len(names)) + "个人哦")
print("**" * 40)
print("宝贝，你要用的底图是" + os.getcwd().replace("\\", "/") + img_path[1:] + "么？")
print("**" * 40)
print("宝贝，你要在" + os.getcwd().replace("\\", "/") + out_path[1:] + "下生成图片么？")
print("**" * 40)
checks_names = input("确认选择Y，否则N，输入一个呗: ")
if checks_names == 'Y' or checks_names == 'y':
    if os.path.exists(out_path) == True:
        backup_path = "backup_" + str(time.mktime(datetime.datetime.now().timetuple()))[:-2]
        shutil.move(out_path, backup_path)
        print("给你保存了一份到" + os.getcwd().replace("\\", "/") + "/" + backup_path + "没用了记得删除哦~~")
        os.mkdir(out_path)
    else:
        os.mkdir(out_path)
    for i in range(3, -1, -1):
        print('\r', '准备kai~ %s 秒！' % str(i).zfill(2), end='')
        time.sleep(1)
    start_time = datetime.datetime.now()
    for name_tmp in names:
        text_width = font.getsize(name_tmp)
        p_name = str(name_tmp + "." + img_type)
        img = Image.open(img_path)
        draw = ImageDraw.Draw(img)
        text_coordinate = int((img.size[0] - text_width[0]) / 2 - x_offset), int(header_y)
        draw.text(text_coordinate, name_tmp, color, font=font)
        img.save(out_path + p_name, img_type, quality=95)
        print('保存成功 在 {}'.format(os.getcwd().replace("\\", "/") + out_path[1:] + p_name))
        numCount = numCount + 1
        print('已经生成 ' + str(numCount) + " 共计", len(names), "还有", str(len(names) - numCount) + " 个要努力干啊")
    end_time = datetime.datetime.now()
    print('亲爱的，搞定啦~，你花了' + str((end_time - start_time))[:-7] + "在" + out_path + "下生成" + str(len(names)) + "张图哦~")
    for i in range(3, -1, -1):
        print('\r', '干完了，回去睡觉了，可以关掉我了，bye，倒数 %s 秒！' % str(i).zfill(2), end='')
        time.sleep(1)

else:
    print("宝贝，再确认一遍吧~excel记得要保存哦。")
