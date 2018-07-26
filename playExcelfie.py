import xlwings as xw
import json

def playFile(excelFileName, jsonFileName):
    with open(jsonFileName) as jsonFile:
        frameList = json.load(jsonFile)

    wb = xw.Book(excelFileName) #set up wb and sheet in xlwings
    ws = wb.sheets[0]

    while True:
        for frame in frameList:

            ws.range((1,1)).value = frame

playFile("excelfieVid.xlsx", "recording_14:12:03.txt")
