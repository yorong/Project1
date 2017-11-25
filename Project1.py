#2017.11.16~
#Project1


import datetime
from selenium import webdriver
import pyautogui
import os
from bs4 import BeautifulSoup
import requests
import sqlite3
import time


class KRX:
    def __init__(self):
        self.driver = webdriver.Firefox()
    #사이트 접속
    def accessToKRX(self):
        self.driver.get('http://kind.krx.co.kr/corpgeneral/corpList.do?method=loadInitPage')
        time.sleep(4)

    #코스피/코스닥 항목 검색
    def search(self, btn_id):
        btn = self.driver.find_element_by_id(btn_id)
        btn.click()
        self.driver.execute_script('fnSearchWithoutIndex();') #검색 버튼 
        time.sleep(3)

    #코스피 엑셀 파일 다운로드
    def downloadKospi(self):
        self.driver.execute_script('fnDownload();')
        time.sleep(4)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(3)
    #코스닥 엑셀 파일 다운로드
    def downloadKosdaq(self):
        self.driver.execute_script('fnDownload();')
        time.sleep(4)
        pyautogui.press('enter')

    '''아 이거 인간적으로 코스피 코스닥 너무 합치고 싶다 ㅠㅠㅠㅠ
     저 다운로드 확인창좀 안나오게 못하나 ㅠㅠ'''


class DBMaker:
    def __init__(self, downLoadPath, dbPath):
        self.downLoadPath = downLoadPath
        self.dbPath = dbPath
        #db 파일 생성 
        self.con = sqlite3.connect(self.dbPath)

    def setData(self, downLoadPath, dbPath):
        self.downLoadPath = downLoadPath
        self.dbPath = dbPath
        self.con = sqlite3.connect(dbPath)

    #db table 생성
    def makeEventTalbe(self, tableName):
        self.cursor = self.con.cursor()
        self.cursor.execute("CREATE TABLE IF NOT EXISTS "+tableName+"(CompanyCode text, Date text, ClosingPrice text, MarketPrice text, HighPrice text, LowPrice text, Volume text)")
    def makeKosTable(self, tableName):
        self.cursor = self.con.cursor()
        self.cursor.execute("CREATE TABLE IF NOT EXISTS "+tableName+"(Date text, CompanyName text, CompanyCode text, Event text, TotalPrice int, Market text, Size text)")


    def closeDB(self):
        self.con.close()


class GetEvents(DBMaker):

    #다운로드 폴더로 이동
    def moveToFileDirection(self):
        os.chdir(self.downLoadPath)

    #xls -> txt 변환
    def changeToTxt(self, fileName):
        if (os.path.isfile(fileName) and fileName == "상장법인목록.xls" ):
            os.rename(fileName, 'KOSPI.txt')
        elif (os.path.isfile(fileName) and fileName == "상장법인목록(1).xls"):
            os.rename(fileName, 'KOSDAQ.txt')

    #파일 열기 
    def openFile(self, fileName):
        self.kospiDoc = open(fileName, "r")
        self.kospiSoup = BeautifulSoup(self.kospiDoc, "html.parser")
    #파일 닫기 
    def closeFile(self):
        self.kospiDoc.close()
    #파일 지우기 
    def deleteFile(self, fileName):
        os.remove(fileName)

    #회사코드 찾기
    def findCompanyCode(self):
        return self.kospiSoup.select('td[style*="@"]')


    #종목별 가격 정보 저장 
    def eventPrices(self, date, tableName, companyCode):
        DBMaker.makeEventTalbe(self, tableName)
        marketTotalPrice= []
        for i in range(10):
            #코드 별 당일 시세 사이트에서 가져오기
            url = "http://finance.naver.com/item/sise_day.nhn?code="+companyCode[i].string
            r = requests.get(url, auth=('user', 'pass'))
            data = r.text
            soup = BeautifulSoup(data, "html.parser")

            #오늘 시세만 추출
            day = soup.find('span', {'class' : 'tah p10 gray03'}).string
            closingPrice = day.find_next().string
            tmp = closingPrice.find_next()
            tmp = tmp.find_next()
            tmp = tmp.find_next()
            marketPrice = tmp.find_next().string
            highPrice = marketPrice.find_next().string
            lowPrice = highPrice.find_next().string
            volume = lowPrice.find_next().string
            #db에 저장
            self.cursor.execute("INSERT INTO "+tableName+" VALUES(?, ?, ?, ?, ?, ?, ?)", (companyCode[i].string, date, closingPrice, marketPrice, highPrice, lowPrice, volume))
            #쉬어가기~
            if (i%10 == 0):
                time.sleep(1)
            #시가 총액 계산
            closingPrice = str(closingPrice).replace(",", "")
            volume = str(volume).replace(",", "")
            marketTotalPrice.append(int(closingPrice)*int(volume))

        #db에 올리기
        self.con.commit()
        return marketTotalPrice



class GetKos(DBMaker):

    #회사 정보 가져와서 저장 
    def companies(self, date, tableName, companyCode, totalPrice, market, size):
        DBMaker.makeKosTable(self, tableName)
        for i in range(10): 
        #save in sqlite
            company = companyCode[i].find_previous_sibling('td').string
            code = companyCode[i].string 
            event = companyCode[i].find_next().string
            self.cursor.execute("INSERT INTO "+tableName+" VALUES(?, ?, ?, ?, ?, ?, ?)", (date, company, code, event, totalPrice[i], market, size[i]))
        self.con.commit()



    #코스피 사이즈 계산
    def kospiSize(self, totalPrice):
        
        kospiSize = [None]*len(totalPrice)
        
        #딕셔너리에 [시가총액:index] 형식으로 저장 
        dic = {}
        for i in range(len(totalPrice)):
            dic[int(totalPrice[i])] = i
        
        #시가총액 내림차순으로 정렬 
        marketRanking = []
        marketRanking = sorted(totalPrice, key=int, reverse=True)
        
        for i in range(len(totalPrice)):
            if (i<=3):
                kospiSize[dic[marketRanking[i]]] = "LargeCap"
            elif (i<=5):
                kospiSize[dic[marketRanking[i]]] = "MidCap"
            else:
                kospiSize[dic[marketRanking[i]]] = "SmallCap"
        
        return kospiSize



    #코스닥 사이즈 계산 
    def kosdaqSize(self, totalPrice):

        kosdaqSize = [None]*len(totalPrice)

        #딕셔너리에 [시가총액:index] 형식으로 저장 
        dic = {}
        for i in range(len(totalPrice)):
            dic[int(totalPrice[i])] = i

        #시가총액 내림차순으로 정렬
        marketRanking = []
        marketRanking = sorted(totalPrice, key=int, reverse=True)
        
        for i in range(len(totalPrice)):
            if (i<=3):
                kosdaqSize[dic[marketRanking[i]]] = "100"
            elif (i<=5):
                kosdaqSize[dic[marketRanking[i]]] = "Mid 300"
            else:
                kosdaqSize[dic[marketRanking[i]]] = "Small"
        
        return kosdaqSize


        


def main():

    #오늘의 날짜
    date = datetime.datetime.now().strftime('%Y%m%d')


    #코스피, 코스닥 excel 파일 다운로드
    krx = KRX()
    krx.accessToKRX()
    krx.search('rWertpapier')
    krx.downloadKospi()
    krx.search('rKosdaq')
    krx.downloadKosdaq()



    #객체 생성
    event = GetEvents("C:\\Users\\SeheeKim\\Downloads", "C:\\Users\\SeheeKim\\Desktop\\Project\\Project1\\Stock.db")
    kos = GetKos("C:\\Users\\SeheeKim\\Downloads", "C:\\Users\\SeheeKim\\Desktop\\Project\\Project1\\Stock.db")
    #다운로드 폴더로 이동 
    event.moveToFileDirection()




    ###코스피
    event.changeToTxt('상장법인목록.xls')
    event.openFile('KOSPI.txt')

    #종목별 코드 찾기 
    kospiCompanyCode = event.findCompanyCode()

    #종목별 정보 저장     
    kospiMarketTotalPrice = event.eventPrices(date, "kospiEvent"+date, kospiCompanyCode)

    #코스피 종목 사이즈 분류 
    kospiSize = kos.kospiSize(kospiMarketTotalPrice)

    #코스피 회사명, 회사코드, 종목 저장 
    kos.companies(date, "kospi"+date, kospiCompanyCode, kospiMarketTotalPrice, "KOSPI", kospiSize)




    ###코스닥
    event.changeToTxt('상장법인목록(1).xls')
    event.openFile('KOSDAQ.txt')

    #종목별 코드 찾기
    kosdaqCompanyCode = event.findCompanyCode()

    #종목별 정보 저장
    kosdaqMarketTotalPrice = event.eventPrices(date, "kosdaqEvent"+date, kosdaqCompanyCode)

    #종목 사이즈 분류
    kosdaqSize = kos.kosdaqSize(kosdaqMarketTotalPrice)

    #회사명, 회사코드, 종목 저장 
    kos.companies(date, "kosdaq"+date, kosdaqCompanyCode, kosdaqMarketTotalPrice, "KOSDAQ", kosdaqSize)



    event.closeDB()
    kos.closeDB()

    #오늘 쓴 파일 지우기
    #event.deleteFile('KOSPI.txt')
    #event.deleteFile('KOSDAQ.txt')






if __name__=="__main__":
    main()



#subprocess
#사이즈 구하는거 범위 바꾸기!! 

#종목별 분류
#실적 
