# pip install openpyxl
# 커맨드가 이거 아닐수도 있어요 어떤 방법이든 일단 openpyxl을 다운받아야 합니다.
import openpyxl as xl

instructions = '''
==========================사용방법==========================
  1. 자신이 번역한 스프레드시트를 .xlsx 포맷으로 다운받으세요.   
  2. 다운로드한 파일의 이름을 subtitles로 바꾸세요. 
  3. subtitles.xlsx를 열어 마지막 줄의 종료 시간을 확인하세요. 
  4. 만약 마지막 줄의 종료 시간이 00:00:00:00이라면 직접 영상을 
     열어 실제 영상의 종료 시간으로 바꿔주세요. 
  5. 파일을 '이 프로그램이 위치한 폴더' 에 넣어주세요.
  6. 모든 준비가 완료되었다면 엔터를 누르세요.
  7. 생성된 자막 파일(jamak.sbv)을 유튜브에 업로드하세요.
==========================주의사항==========================
  1. 번역을 진행한 엑셀 파일의 이름은 반드시 subtitles.xlsx
     이어야 하며, 위치 역시 이 프로그램의 위치와 같아야 합니다.
  2. 생성된 자막 파일 역시 같은 위치에 생성됩니다.
  3. 만약 마지막 줄의 종료 시간을 수정하지 않았다면 유튜브에 
     업로드하고 직접 수정해줘야 합니다.
===========================================================
'''

cnt = 0

time_list = []
sub_list = []
# -*- coding: utf-8 -*-
import sys
import re
from time import sleep

print(instructions)
cnt = 0
while True:
    text = input("준비되셨다면 엔터를 누르세요!")
    if text == "":
        break
    else:
        print("%s 말고 엔터를 누르세요" % (text))
        cnt += 1
        continue


out = open("jamak.sbv", "w", encoding='utf8')
sub_ = xl.load_workbook("subtitles.xlsx")
sub = sub_.active
row_cnt = sub.max_row
rows = list(sub.rows)


for i in range(5, row_cnt, 2):
    r = rows[i]
    if r[0].value == None:
        break
    start_time = r[1].value
    end_time = r[2].value

    start_time = start_time[:8] + '.' + start_time[9:]
    end_time = end_time[:8] + '.' + end_time[9:]

    time_list.append(start_time + '0,' + end_time + '0\n')


for i in range(6, row_cnt, 2):
    r = rows[i]
    if r[3].value == None:
        break
    res = r[4].value
    if res == None:
        res = ""
    sub_list.append(res)


cnt = 0
for time in time_list:
    out.write(time)
    try:
        out.write(str(sub_list[cnt]) + "\n\n")
    except IndexError as e:
        break
    cnt += 1

out.close()
print("\n\n변환 완료!")
