#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
#Created on Wed Oct 27 18:21:57 2021

#@author: ljjin9711

#엑셀 파일 확장자는 xls만 지원함
"""

import xlrd
import xlwt
import random
import tkinter as tk #GUI

''' 엑셀 셀 스타일 '''
ValueStyle = xlwt.easyxf(
        'border: top thin, right thin, bottom thin, left thin; align: horizontal center; font: height 250')
NumberStyle = xlwt.easyxf(
        'border: top thin, right thin, bottom thin, left thin; align: horizontal center; font: height 440')
DateStyle = xlwt.easyxf('font: height 200')
TitleStyle = xlwt.easyxf('font: height 250')
subTitleStyle = xlwt.easyxf('font: height 0')
LangStyle = xlwt.easyxf(
        'border: top thin, right thin, bottom thin, left thin; pattern: pattern solid, fore_color gray25; align: horizontal center; font: height 440')

stateText = "** 파일 경로를 제대로 작성해주세요!"

def Main():
    WordFileArea = WordOldNameEntry.get() #단어 파일 경로 및 이름
    WordCreateFileArea = WordNewNameEntry.get() #저장할 단어 파일 경로와 이름
    wb = xlrd.open_workbook(WordFileArea)  # 파일 읽기
    sheets = wb.sheets() 
    mode = Modelistbox.curselection()[0] #단어 모드
    WordCount = int(WordCountEntry.get()) #단어 개수
    wbwt = xlwt.Workbook(encoding='utf-8')
    ws = []
    if mode == 1:
        createKoreanRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt)
    elif mode == 2:
        createEnglishRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt)
    elif mode == 3:
        createFirstEnglishRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt)
    else: #섞어섞어
         createMixRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt)


def baseSetting(ws, sheetNumber): #시트마다 앞에 생성될 셀들
    ws[sheetNumber].write(0, 0, "20__년 __월 __일", DateStyle)
    ws[sheetNumber].write(1, 0, "20__학년도 __학기 어휘 평가   학번_______ 이름_______ 점수______", TitleStyle)
    ws[sheetNumber].write(2, 0, " ", subTitleStyle)
    ws[sheetNumber].write(3, 1, 'English', LangStyle)
    ws[sheetNumber].write(3, 2, 'Korean', LangStyle)
    ws[sheetNumber].write(3, 0, "", ValueStyle) 
    ws[sheetNumber].write(3, 4, 'English', LangStyle)
    ws[sheetNumber].write(3, 5, 'Korean', LangStyle)
    ws[sheetNumber].write(3, 3, "", ValueStyle) 

def cellSetting(ws, sheetNumber): #시트들의 간격과 크기를 조정
    ws[sheetNumber].col(0).width = 10 * 185
    ws[sheetNumber].col(1).width = 40 * 122
    ws[sheetNumber].col(2).width = 40 * 122
    ws[sheetNumber].col(3).width = 10 * 185
    ws[sheetNumber].col(4).width = 40 * 122
    ws[sheetNumber].col(5).width = 40 * 122

def randValue(start, end):
    WordValue = list(range(start, end))
    random.shuffle(WordValue)
    return WordValue

def createEnglishRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt): #영어 랜덤 
    start = 4; #단어 시작
    isChange = True #양식의 좌우 전환
    
    sheetNumber = (WordCount // 50) + 1 #총 시트 수

    WordNumber = 5 #단어 번호
    WordSuppoter = -1 #단어 번호 도우미

    WordValue = randValue(0, WordCount)

    ws = [0 for k in range(sheetNumber)] #시트를 배열 화
    for n in range(0, sheetNumber-1, 1): #시트를 하나씩 셍성
        ws[n] = wbwt.add_sheet(str(n), cell_overwrite_ok=True) 

    for i in range(0, sheetNumber-1, 1): #시트 수 만큼 반복
        WordSuppoter += 1

        baseSetting(ws, i); 

        for num in range(start, 50+start, 1):
            WordNumber = (50 * WordSuppoter) + (num - 3) #단어 - 스타트
            
            value = sheets[0].cell_value(WordValue[WordNumber-1], 1) #이 값을 바꾸면서 사용
            eng_value = sheets[0].cell_value(WordValue[WordNumber-1], 0) #이 값을 바꾸면서 사용
            if isChange == False:
                ws[i].write(num - 25, 4, eng_value, ValueStyle)  # 단어
                ws[i].write(num - 25, 3, WordNumber, NumberStyle)  # 번호
                ws[i].write(num - 25, 5, "", ValueStyle)
            else:
                ws[i].write(num, 1, eng_value, ValueStyle)  # 단어
                ws[i].write(num, 0, WordNumber, NumberStyle)  # 번호
                ws[i].write(num, 2, "", ValueStyle)
            

            if WordNumber % 25 == 0:
                if isChange == True:
                    isChange = False
                else:
                    isChange = True

            cellSetting(ws, i)

    wbwt.save(WordCreateFileArea) #파일 저장
    global stateText
    stateText = "성공적으로 단어 파일을 만들었습니다."
    stateLabel.config(text=stateText)
    return


def createKoreanRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt): #한글 랜덤 
    start = 4; #단어 시작
    isChange = True #양식의 좌우 전환
    
    sheetNumber = (WordCount // 50) + 1 #총 시트 수

    WordNumber = 5 #단어 번호
    WordSuppoter = -1 #단어 번호 도우미

    WordValue = randValue(0, WordCount)

    ws = [0 for k in range(sheetNumber)] #시트를 배열 화
    for n in range(0, sheetNumber-1, 1): #시트를 하나씩 셍성
        ws[n] = wbwt.add_sheet(str(n), cell_overwrite_ok=True) 

    for i in range(0, sheetNumber-1, 1): #시트 수 만큼 반복
        WordSuppoter += 1

        baseSetting(ws, i); 

        for num in range(start, 50+start, 1):
            WordNumber = (50 * WordSuppoter) + (num - 3) #단어 - 스타트
            
            value = sheets[0].cell_value(WordValue[WordNumber-1], 1) #이 값을 바꾸면서 사용
            eng_value = sheets[0].cell_value(WordValue[WordNumber-1], 0) #이 값을 바꾸면서 사용

            if isChange == False:
                ws[i].write(num - 25, 4, "", ValueStyle)  # 단어
                ws[i].write(num - 25, 3, WordNumber, NumberStyle)  # 번호
                ws[i].write(num - 25, 5, value, ValueStyle)
            else:
                ws[i].write(num, 1, "", ValueStyle)  # 단어
                ws[i].write(num, 0, WordNumber, NumberStyle)  # 번호
                ws[i].write(num, 2, value, ValueStyle)
            

            if WordNumber % 25 == 0:
                if isChange == True:
                    isChange = False
                else:
                    isChange = True

            cellSetting(ws, i)

    wbwt.save(WordCreateFileArea) #파일 저장
    global stateText
    stateText = "성공적으로 단어 파일을 만들었습니다."
    stateLabel.config(text=stateText)
    return

def createMixRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt): #한글, 영어 랜덤 
    start = 4; #단어 시작
    isChange = True #양식의 좌우 전환
    
    sheetNumber = (WordCount // 50) + 1 #총 시트 수

    WordNumber = 5 #단어 번호
    WordSuppoter = -1 #단어 번호 도우미

    WordValue = randValue(0, WordCount)


    ws = [0 for k in range(sheetNumber)] #시트를 배열 화
    for n in range(0, sheetNumber-1, 1): #시트를 하나씩 셍성
        ws[n] = wbwt.add_sheet(str(n), cell_overwrite_ok=True) 

    for i in range(0, sheetNumber-1, 1): #시트 수 만큼 반복
        WordSuppoter += 1

        baseSetting(ws, i); 

        for num in range(start, 50+start, 1):
            WordNumber = (50 * WordSuppoter) + (num - 3) #단어 - 스타트
            
            randMix = random.randint(0, 1) #영어셀 한글셀 랜덤

            if randMix == 0:

                value = sheets[0].cell_value(WordValue[WordNumber-1], 1) #이 값을 바꾸면서 사용
                eng_value = sheets[0].cell_value(WordValue[WordNumber-1], 0) #이 값을 바꾸면서 사용

                if isChange == False:
                    ws[i].write(num - 25, 4, "", ValueStyle)  # 단어
                    ws[i].write(num - 25, 3, WordNumber, NumberStyle)  # 번호
                    ws[i].write(num - 25, 5, value, ValueStyle)
                else:
                    ws[i].write(num, 1, "", ValueStyle)  # 단어
                    ws[i].write(num, 0, WordNumber, NumberStyle)  # 번호
                    ws[i].write(num, 2, value, ValueStyle)
            else:
                value = sheets[0].cell_value(WordValue[WordNumber-1], 0) #이 값을 바꾸면서 사용
                if isChange == False:
                    ws[i].write(num - 25, 4, value, ValueStyle)  # 단어
                    ws[i].write(num - 25, 3, WordNumber, NumberStyle)  # 번호
                    ws[i].write(num - 25, 5, "", ValueStyle)
                else:
                    ws[i].write(num, 1, value, ValueStyle)  # 단어
                    ws[i].write(num, 0, WordNumber, NumberStyle)  # 번호
                    ws[i].write(num, 2, "", ValueStyle)

            if WordNumber % 25 == 0:
                if isChange == True:
                    isChange = False
                else:
                    isChange = True

            cellSetting(ws, i)

    wbwt.save(WordCreateFileArea) #파일 저장
    global stateText
    stateText = "성공적으로 단어 파일을 만들었습니다."
    stateLabel.config(text=stateText)
    return

def createFirstEnglishRandomFile(WordCreateFileArea, WordCount, sheets, ws, wbwt): #영어 첫글자 한글 랜덤 
    start = 4; #단어 시작
    isChange = True #양식의 좌우 전환
    
    sheetNumber = (WordCount // 50) + 1 #총 시트 수

    WordNumber = 5 #단어 번호
    WordSuppoter = -1 #단어 번호 도우미
    FirstWord = ""
    WordValue = randValue(0, WordCount)

    ws = [0 for k in range(sheetNumber)] #시트를 배열 화
    for n in range(0, sheetNumber-1, 1): #시트를 하나씩 셍성
        ws[n] = wbwt.add_sheet(str(n), cell_overwrite_ok=True) 

    for i in range(0, sheetNumber-1, 1): #시트 수 만큼 반복
        WordSuppoter += 1

        baseSetting(ws, i); 

        for num in range(start, 50+start, 1):
            WordNumber = (50 * WordSuppoter) + (num - 3) #단어 - 스타트
            
            value = sheets[0].cell_value(WordValue[WordNumber-1], 1) #이 값을 바꾸면서 사용
            eng_value = sheets[0].cell_value(WordValue[WordNumber-1], 0) #이 값을 바꾸면서 사용
            for l in eng_value:
                FirstWord = l
                break
            if isChange == False:
                ws[i].write(num - 25, 4, FirstWord+"             ", ValueStyle)  # 단어
                ws[i].write(num - 25, 3, WordNumber, NumberStyle)  # 번호
                ws[i].write(num - 25, 5, value, ValueStyle)
            else: 
                ws[i].write(num, 1, FirstWord+"             ", ValueStyle)  # 단어
                ws[i].write(num, 0, WordNumber, NumberStyle)  # 번호
                ws[i].write(num, 2, value, ValueStyle)
            

            if WordNumber % 25 == 0:
                if isChange == True:
                    isChange = False
                else:
                    isChange = True

            cellSetting(ws, i)

    wbwt.save(WordCreateFileArea) #파일 저장
    global stateText
    stateText = "성공적으로 단어 파일을 만들었습니다."
    stateLabel.config(text=stateText)
    return

root = tk.Tk() #가장 상위 레벨 창 생성
root.title("Dlmajang - AutoWordTest") #창 제목
root.geometry("640x350") #창 크기
root.resizable(False, False) #창 크기 조절 여부

emptyLabel = tk.Label(root, text="", width=50)
emptyLabel.grid(row=8, column=2)

#========
# 제목
Titlelabel = tk.Label(root, text="환영합니다. 아래 보이는 빈칸을 모두 채워주셔야 합니다. (xxx.xls만 지원)\n문의 : dlmajang@naver.com", width=50, height=3, fg="black", relief="sunken")
Titlelabel.grid(row=0, column=2)
#======== 
# 불러올 단어파일
WordOldNameEntryLabel = tk.Label(root, text="Old 파일 경로 :", width=10)
WordOldNameEntryLabel.grid(row=1, column=0)

WordOldNameEntry=tk.Entry(root, width=50, relief="sunken")
WordOldNameEntry.insert(0, "저장되어 있는 단어 파일 경로(이름과 확장자 포함)")
WordOldNameEntry.grid(row=1, column=2)
#======== 
# 만들 단어 파일
WordNewNameEntryLabel = tk.Label(root, text="New 파일 경로 :", width=10)
WordNewNameEntryLabel.grid(row=2, column=0)

WordNewNameEntry=tk.Entry(root, width=50, relief="sunken")
WordNewNameEntry.insert(0, "생성시킬 단어 파일 경로(이름과 확장자 포함)")
WordNewNameEntry.grid(row=2, column=2)
#========
# 단어 시험지 모드
ModelistboxLabel = tk.Label(root, text="Mode", width=10)
ModelistboxLabel.grid(row=5, column=0)

Modelistbox = tk.Listbox(root, selectmode='browse', width=50, height=0)
Modelistbox.insert(0, "========선택========")
Modelistbox.insert(1, "1. 영어빈칸 + 한글해석")
Modelistbox.insert(2, "2. 영어해석 + 한글빈칸")
Modelistbox.insert(3, "3. 첫글자 영어 + 한글해석")
Modelistbox.insert(4, "4. 1+2 섞어서")
Modelistbox.grid(row=5, column=2)
#========
# 단어 개수
WordCountEntryLabel = tk.Label(root, text="단어 개수\n(50배수)", width=0)
WordCountEntryLabel.grid(row=6, column=0)

WordCountEntry = tk.Entry(root, width=50, relief="sunken")
WordCountEntry.insert(0, "단어 개수(50배수)")
WordCountEntry.grid(row=6, column=2)
#========
# 상태 알림
stateLabel = tk.Label(root, text=stateText, width=50)
stateLabel.grid(row=7, column=2)

ConvertButton = tk.Button(root, text="Convert", overrelief="solid", width=15, command=Main, repeatdelay=1000, repeatinterval=100)
ConvertButton.grid(row=10, column=2)

root.mainloop() #창이 종료시까지 반복
