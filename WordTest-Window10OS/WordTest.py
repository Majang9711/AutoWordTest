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


def Main():
    print("=== 자동 단어시험지 생성기 ===")
    WordFileArea = input("사용할 단어파일의 경로를 입력해주세요 > ")
    WordCreateFileArea = input("생성될 단어 시험지 파일의 이름을 입력해주세요 > ")
    wb = xlrd.open_workbook(WordFileArea)  # 파일 읽기
    sheets = wb.sheets()
    mode = int(input("=== 출력할 모드를 선택해주세요 === \n1. 영어빈칸 + 한글해석 \n2. 영어해석 + 한글빈캄 \n3. 첫글자 영어 + 한글해석\n4. 1+2 섞어서\n해당 번호를 입력해주세요 > "))
    WordCount = int(input("단어 개수(50의 배수로만 입력가능) > "))
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
    ws[sheetNumber].col(0).width = 10 * 300
    ws[sheetNumber].col(1).width = 40 * 122
    ws[sheetNumber].col(2).width = 40 * 122
    ws[sheetNumber].col(3).width = 10 * 300
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

    return


if __name__ == '__main__':
    Main() #Main 함수 실행