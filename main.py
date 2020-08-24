# -*- coding:utf-8 -*-
import importlib, sys
import configparser
import xlrd
import chardet
importlib.reload(sys)
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
iniconfig = configparser.ConfigParser()
fd =  open('./config.ini', "rb")
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

for name_tmp in names:
    text_width = font.getsize(name_tmp)
    p_name = str(name_tmp + '.jpg')
    img = Image.open(img_path)
    draw = ImageDraw.Draw(img)
    text_coordinate = int((img.size[0] - text_width[0]) / 2 - x_offset), int(header_y)
    draw.text(text_coordinate, name_tmp, color, font=font)
    img.save('./output/' + p_name, 'jpeg',quality=95)
    print('保存成功 at {}'.format('./output/' + p_name, 'png'))
    numCount = numCount + 1
    print('完成 ' + str(numCount) + ' 共计', len(names), '剩余 ' ,(len(names)-numCount))

