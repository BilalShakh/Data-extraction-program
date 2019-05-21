import pyautogui
import clipboard
from openpyxl import Workbook
from openpyxl import load_workbook


#put in the path to workbook
mainWorkbook = load_workbook(r'C:\Users\Shazia\Documents\Clare data and research\house prices\Data 23,04,19.xlsx')
ActiveWorkbook = mainWorkbook.active

def convertMonthDay(Date):
    Days = 0
    Months = 0
    endMonIndex = 0
    starDayIndex = 0
    
    for i in range(len(Date)):
        if (Date[i] == " "):
            endDateIndex = i
            break

    Months = int(Date[0:endDateIndex])

    for i in range(-1,-len(Date),-1):
        if (Date[i] == " "):
            starDayIndex = i+1
            break

    Days = int(Date[starDayIndex:]) + Months*30

    return(Days)

def getBBWP(DataStr):
    endBedIndex = 0
    endBaseIndex = 1
    endWashIndex = 0
    endParkIndex = 0
    baseNo = 0


    for i in range(len(DataStr)):
        if(not str.isdigit(DataStr[i])):
            endBedIndex = i
            break

    for i in range(endBedIndex+1, len(DataStr)):
        if(not str.isdigit(DataStr[i])):
            endBaseIndex = i
            break


    for i in range(endBaseIndex+4, len(DataStr)):
        if(not str.isdigit(DataStr[i])):
            endWashIndex = i
            break

    for i in range(endWashIndex+4, len(DataStr)):
        if(not str.isdigit(DataStr[i])):
            endParkIndex = i
            break

    try:
        baseNo = int(DataStr[endBedIndex+1:endBaseIndex])
    except:
        baseNo = 0

    bedNo = int(DataStr[0:endBedIndex])

    if(baseNo==0):
        washNo = int(DataStr[endBaseIndex+3:endWashIndex])
    else:
        washNo = int(DataStr[endBaseIndex+4:endWashIndex])

    parkNo = int(DataStr[endWashIndex+4:endParkIndex])

    return(bedNo,baseNo,washNo,parkNo)
    



for q in range(8):
    y = 265
    for i in range(11):
        sellPrice, noOfBeds, noOfWashrooms, noOfPark, daysOnMarket, noOfBase = 0,0,0,0,0,0
        Neighbourhood = ""
    
        pyautogui.click(1230, y, clicks=1, interval=0, button='left')
    
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')

        AllText = clipboard.paste()
        AlltextList = AllText.splitlines()

        
        for q in AlltextList:

            if ("Sold -" in q):
                sellPrice = int(q.replace("Sold - $","").replace(",",""))

            elif ("  | " in q and str.isdigit(q[0])):
                noOfBeds, noOfBase, noOfWashrooms, noOfPark= getBBWP(q)

            elif ("Days on market:" in q):
                Days = q.replace("Days on market: ","").replace(" days","")

                if not (str.isdigit(Days)):
                    daysOnMarket = convertMonthDay(Days)
                else:
                    daysOnMarket = int(Days)

            elif ("Neighbourhood: " in q):
                Neighbourhood = q.replace("Neighbourhood: ","")
            
    
        ActiveWorkbook.append([sellPrice, noOfBeds, noOfWashrooms, noOfPark, daysOnMarket, Neighbourhood, noOfBase])       
    
        pyautogui.click(1668, 163, clicks=1, interval=0, button='left')
        y += 65
    pyautogui.moveTo(1600, 472,0)
    pyautogui.scroll(-575)


    
mainWorkbook.save(r'C:\Users\Shazia\Documents\Clare data and research\house prices\Data 23,04,19.xlsx')

    
    #(x=1526, y=161)
    #(x=500, y=50)
    #pyautogui.moveTo(1230, x, duration=0)
    #pyautogui.scroll(-960)


