# -*- coding:utf-8 -*-
import importlib, sys, os, datetime, shutil, time
import configparser
import xlrd
import chardet
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
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
importlib.reload(sys)
print("**" * 40)
print("**" * 40)
print("**" * 40)
print("宝贝，稍等一下，给你看个预览。")

for na_tmp in names[0:1]:
    text_width = font.getsize(na_tmp)
    img = Image.open(img_path)
    draw = ImageDraw.Draw(img)
    text_coordinate = int((img.size[0] - text_width[0]) / 2 - x_offset), int(header_y)
    draw.text(text_coordinate, na_tmp, color, font=font)
    img.show()