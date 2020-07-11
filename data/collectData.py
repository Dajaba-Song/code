from selenium import webdriver
import time
import openpyxl
import re

from urllib.request import urlretrieve
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

# 텍스트에 포함되어 있는 특수 문자 제거
def cleanText(movie):
    text = re.sub('[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\‘|\(\)\[\]\<\>`\'…》]', '', movie)
    return text

wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(["영화", "감독", "배우", "코드번호", "장르", "상영타입", "매출액", "관객 수"])

# options = webdriver.ChromeOptions()
# options.add_argument("headless")
# driver = webdriver.Chrome("./chromedriver", options=options)

driver = webdriver.Chrome("./chromedriver")
driver.get("http://www.kobis.or.kr/kobis/business/stat/boxs/findYearlyBoxOfficeList.do")
time.sleep(3)

movieList = []

# 2004 ~ 2020년까지 연도별로
for year in range(2004, 2021) :
    time.sleep(2)
    yearBtn = driver.find_element_by_xpath("//select[@name='sSearchYearFrom']/option[@value='" + str(year) + "']")
    yearBtn.click()
    clickBtn = driver.find_element_by_css_selector("button.btn_blue")
    clickBtn.click()

    # 각각 50개의 영화 데이터 수집
    time.sleep(2)
    for num in range(0, 50) :
        container = driver.find_element_by_css_selector("tr#tr_"+str(num))
        link = container.find_element_by_css_selector("a")
        driver.execute_script("arguments[0].click();", link)

        information = []

        # 영화 제목
        tmp = driver.find_elements_by_css_selector("strong.tit")
        movie = tmp[len(tmp) - 1].text.strip()
        movie = cleanText(movie)

        information.append(movie)
        movieList.append(movie)

        movieDriver = webdriver.Chrome("./chromedriver.exe")
        movieDriver.get("https://search.daum.net/search?q="+movie)

        # 감독
        directors = ""
        directorList = movieDriver.find_elements_by_css_selector("dl:nth-of-type(2) > dd.cont > a.stit")
        for director in directorList :
            directors += director.text.strip()
            directors += ","
        directors = directors[:-1]
        information.append(directors)

        # 배우
        actors = ""
        actorList = movieDriver.find_elements_by_css_selector("dl:nth-of-type(3) > dd.cont > a.stit")
        for actor in actorList:
            actors += actor.text.strip()
            actors += ","
        actors = actors[:-1]
        information.append(actors)

        movieDriver.close()

        # 영화코드, 장르, 상영타입
        for n in range(1, 11) :
            if (n == 1 or n == 4 or n == 10) :
                info = driver.find_element_by_css_selector("dl.ovf dd:nth-of-type(" + str(n) + ")").text.strip()
                if (n == 1) :
                    imageName = info
                    info = int(info)
                if (n == 4) :
                    info = info.split('|')[2].strip()
                information.append(info)

        # 포스터
        imageUrl = driver.find_element_by_css_selector("a.fl.thumb").get_attribute("href")
        urlretrieve(imageUrl, 'image/'+imageName+'.jpg')

        closeBtn = driver.find_element_by_css_selector("div.hd_layer a:nth-of-type(2)")
        driver.execute_script("arguments[0].click();", closeBtn)

        # 매출액
        earn = container.find_element_by_css_selector("td#td_salesAcc").text.strip()
        earn = earn.replace(",", "")
        earn = int(earn)
        information.append(earn)

        # 관객 수
        audience = container.find_element_by_css_selector("td#td_audiAcc").text.strip()
        audience = audience.replace(",", "")
        audience = int(audience)
        information.append(audience)

        sheet.append(information)

wb.save("MovieData.xlsx")
driver.close()