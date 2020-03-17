# IngressKeyGet GetKey.py
# 2019/9/12 19:36

__author__ = 'CHAZHOU <konglc123@outlook.com>'

import cv2
from aip import AipOcr
import xlwt
import shutil
import os

InPath = "D:/IngressKeyGet/keypic/"
OutPathName = "D:/IngressKeyGet/keyname/"
OutPathCount = "D:/IngressKeyGet/keycount/"

shutil.rmtree(InPath)
shutil.rmtree(OutPathCount)
shutil.rmtree(OutPathName)

os.makedirs(InPath)
os.makedirs(OutPathName)
os.makedirs(OutPathCount)

# 视频转换成图片
vc = cv2.VideoCapture('D:\IngressKeyGet\key.mp4')#这是我的路径
rval = vc.isOpened()
c=0
while rval:
    rval, frame = vc.read()#farame就是帧对象
    cv2.imwrite('D:\IngressKeyGet\keypic/'+str(c) + '.jpg', frame) #i
    c=c+1
vc.release()


#图片缩小范围
name_xmin = 100
name_ymin = 560
name_width = 340
name_height = 40
count_xmin = 380
count_ymin = 320
count_width = 50
count_height = 50

for num in range(0, c-1):
    img = cv2.imread(InPath+str(num)+'.jpg')
    cropImgname = img[name_ymin:name_ymin+name_height,name_xmin:name_xmin+name_width]
    # 获取需要部分的图像
    cropImgname = cv2.cvtColor(cropImgname,cv2.COLOR_BGR2GRAY)
    cv2.imwrite(OutPathName+str(num)+'.jpg',cropImgname)
    cropImgcount = img[count_ymin:count_ymin+count_height,count_xmin:count_xmin+count_width]
    cropImgcount = cv2.cvtColor(cropImgcount,cv2.COLOR_BGR2GRAY)
    cv2.imwrite(OutPathCount+str(num)+'.jpg',cropImgcount)
# img = cv2.imread('D:/IngressKeyGet/keypic/27.jpg')
# cv2.namedWindow('image')
# cv2.rectangle(img,(name_xmin,name_ymin),(name_xmin+name_width,name_ymin+name_height),(0,255,0),2)
# cv2.rectangle(img,(count_xmin,count_ymin),(count_xmin+count_width,count_ymin+count_height),(0,255,0),2)
# cv2.imshow("image",img)
# cv2.waitKey(0)
# cv2.destroyAllWindows()


#百度ai

APP_ID = 'MyAPPID'
API_KEY = 'MyapiKey'
SECRET_KEY = 'MyapiSecretKey'
client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

fname = '/IngressKeyGet/keyname/'
fcount = '/IngressKeyGet/keycount/'

book = xlwt.Workbook()
sheet = book.add_sheet('ingresskeycount')

# 读取文件
def get_file_content(filepath):
    with open(filepath,'rb') as fp:
        return fp.read()

for num in range(0,c-1):
    nameimage = get_file_content(fname+str(num)+'.jpg')
    countimage = get_file_content(fcount+str(num)+'.jpg')

    # 调用识别
    try:
        nameresults = client.general(nameimage)["words_result"]
        countresult = client.general(countimage)["words_result"]
    except KeyError:
        continue
    nameimg = cv2.imread(fname)
    countimg = cv2.imread(fcount)


    for result in nameresults:
        name = result["words"]
        sheet.write(num, 0, name)
        book.save('\IngressKeyGet\keycount.xls')
        # namelocation = result["location"]
        # print(name)

    for result in countresult:
        text = result["words"]
        # print(text)
        count = list(filter(str.isdigit,text))
        number = str(list(map(int,count)))
        try:
            number = int(number[1:len(number)-1]) # 取数字
        except ValueError:
            continue
        sheet.write(num, 1, number)
        book.save('\IngressKeyGet\keycount.xls')

print('successful')
    # print(number)
    # print(type(number))
    # #rectangele画框
    # cv2.rectangle(nameimg,(namelocation["left"],namelocation["top"]),
    #               (namelocation["left"]+namelocation["width"],namelocation["top"]+namelocation["height"]),
    #               (0,255,0),2)
    #
    # cv2.imwrite(fname[:-4] + "_result.jpg", nameimg)
