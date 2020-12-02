# coding=utf-8
import configparser

import chardet
import cv2
import numpy
import xlrd
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
table = data.sheet_by_index(0)
names = table.col_values(0)[1:]
data.sheet_names()
inputfont = iniconfig.get("source", "inputfont")
fontsize = int(iniconfig.get("image", "fontsize"))
font = ImageFont.truetype(inputfont, fontsize)
x_offset = int(iniconfig.get("image", "x_offset"))
header_y = iniconfig.get("image", "header_y")
video_path = iniconfig.get("source", "video_path")
out_path = iniconfig.get("image", "out_path")
color = iniconfig.get("image", "color")


def cv2ImgAddText(img, na_tmp):
    img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(img)
    text_width = font.getsize(na_tmp)
    font.getsize(na_tmp)
    text_coordinate = int((img.size[0] - text_width[0]) / 2 - x_offset), int(header_y)
    draw.text(text_coordinate, na_tmp, font=font, fill=color)
    return cv2.cvtColor(numpy.asarray(img), cv2.COLOR_RGB2BGR)


for na_tmp in names:
    print(na_tmp)
    cap = cv2.VideoCapture(video_path)
    fps_video = cap.get(cv2.CAP_PROP_FPS)
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    frame_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    frame_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    videoWriter = cv2.VideoWriter(out_path + na_tmp + '.mp4', fourcc, fps_video, (frame_width, frame_height))
    success = True
    while success:
        success, frame = cap.read()
        if not success:
            break
        frame = cv2ImgAddText(frame, na_tmp)
        # cv2.imshow("Image", frame)
        videoWriter.write(frame)
    #  if cv2.waitKey(1) == ord('q'):
    #     break
    cap.release()
# cv2.destroyAllWindows()
