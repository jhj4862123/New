import re
import os
import shutil
import pyautogui
import openpyxl
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import pandas as pd
from datetime import datetime
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tqdm import tqdm

############################ 이미지 폴더 선택 ########################################
root = Tk()
root.title("폴더 선택 창")   # 타이틀 설정

file_frame = Frame(root)
file_frame.pack(fill="x", padx = 5, pady= 5)

root.geometry("320x240") # 가로 *세로 사이즈
root.resizable(False, False)    #가로 *세로 사이즈 변경 가능 유무

dir_path = None        #폴더 경로 담을 변수 생성

def folder_select():
    global dir_path
    dir_path = filedialog.askdirectory(initialdir="./", \
                                       title = "폴더를 선택 해 주세요")  #folder 변수에 선택 폴더 경로 넣기
    if dir_path == '':
        messagebox.showwarning("경고", "폴더를 선택 하세요")    #폴더 선택 안했을 때 메세지 출력
    else:
        res = os.listdir(dir_path) # 폴더에 있는 파일 리스트 넣기
        if len(res) == 0:
            messagebox.showwarning("경고", "폴더내 파일이 없습니다.")
        else:
            root.destroy()

btn_active_dir = Button(file_frame, text ="충전기 사진을 선택해 주세요. \n\n사진 형식 : 충전기번호_1.jpg\n ex) 1234_1.jpg", \
                        font=36, width = 24, padx = 10, pady= 20, command=folder_select)
btn_active_dir.pack( padx = 5, pady= 5)

root.mainloop()

############################ 경로 및 양식 ########################################

now = datetime.now()
s = now.strftime("%Y-%m-%d")
finishpath = '완료폴더/'
newpath = finishpath + s

# photosrc = '작업 전 사진/'
photosrc = dir_path + '/'
movephoto = newpath + '/완료된 사진/'
move_resize_photo = newpath + '/축소 사진/'
movefilesrc = '완료폴더/'
path = '점검데이터.xlsx'
j = 1

print("\nphotosrc : ", photosrc)

if not os.path.exists(newpath):
    os.makedirs(newpath)

if not os.path.exists(movephoto):
    os.makedirs(movephoto)

if not os.path.exists(move_resize_photo):
    os.makedirs(move_resize_photo)

data = pd.read_excel(path)
base = photosrc
print("\nbase : ", base)

count_photo = [] # 사진의 갯수

############################ 파일 분리 ############################################################
file_names = []

file_names = os.listdir(dir_path)
print(f"file_names : {file_names}")
for name in file_names :
    src_name = name
    temp_name = re.split('[,|_|.]', name)
    print(f"글자수 : {len(temp_name)}")
    if(len(temp_name) != 3) :

        print(f"temp_name : {temp_name}")

        for j in range(0, len(temp_name) - 2) :
            print(f"글자 분리 : {temp_name[j]}")
            src = os.path.join(photosrc, name)
            print(f"src : {src}")
            dst = temp_name[j] + '_' + temp_name[-2] + '.jpg'
            dst = os.path.join(photosrc, dst)
            print(f"dst : {dst}")
            shutil.copyfile(src, dst)