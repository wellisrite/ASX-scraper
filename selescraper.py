from selenium import webdriver
from selenium import common
from selenium.webdriver.chrome.options import Options
from time import sleep
import csv


def requestadjustment(i):
    print "Fetching %s data \n ------------------------------------" % i
    # price
    try:
        price = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[1]/td[1]/span').text
        with open('asxprices.csv', mode='a') as doc:
            writer = csv.writer(doc, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, price
            writer.writerow(string)
        # marketcap
        marketcap = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[2]/td[1]/div/span').text
        with open('asxmarketcaps.csv', mode='a') as doc:
            writer = csv.writer(doc, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, marketcap
            writer.writerow(string)
        # dividend
        mostrecent = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[2]/span[1]').text
        exdate = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]').text
        paydate = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[3]/td[2]').text
        franking = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[4]/td[2]').text
        print "price :", price, '\n', "marketcap :", marketcap, '\n', "pdfurl:", pdfurl, '\n'",mostrecent :", mostrecent, '\n', "exdate :", exdate, '\n', "paydate :", paydate, '\n', "franking :", franking
        with open('asxdividens.csv', mode='a') as doc:
            writer = csv.writer(doc, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, mostrecent, exdate, paydate, franking
            writer.writerow(string)

        # key statistics
        try:
            driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[3]/ul/li[2]/a').click()
            driver.implicitly_wait(5)
            # #day
            openprice = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[3]/td[2]/span').text
            highprice = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[4]/td[2]/span').text
            lowprice = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[5]/td[2]/span').text
            volume = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[6]/td[2]/span').text
            bid = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[7]/td[2]/span').text
            offer = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[8]/td[2]/span').text
            numerofshares = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[9]/td[2]/span').text
            # year
            prevclose = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[3]/td[4]/span').text
            week52h = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[4]/td[4]/span').text
            week52l = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[5]/td[4]/span').text
            averagevol = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[6]/td[4]/span').text
            # ratios
            pore = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[3]/td[6]/span').text
            eps = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[4]/td[6]/span').text
            ady = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[5]/td[6]/span').text
            keystatisticsasx.append(i+","+openprice+","+highprice+","+lowprice+","+volume+","+bid+","+offer +
                                    ","+numerofshares+","+prevclose+","+week52h+","+week52l+","+averagevol+","+pore+","+eps+","+ady)
            print openprice, "\n", highprice, "\n", lowprice, "\n", volume, "\n", bid, "\n", offer, "\n", numerofshares, "\n", prevclose, "\n", week52h, "\n", week52l, "\n", averagevol, "\n", pore, "\n", eps, "\n", ady
            with open('asxkeystatistics.csv', mode='a') as doc:
                writer = csv.writer(doc, delimiter=",",
                                    quotechar='"', quoting=csv.QUOTE_MINIMAL)
                string = i, openprice, highprice, lowprice, volume, bid, offer, numerofshares, prevclose, week52h, week52l, averagevol, pore, eps, ady
                writer.writerow(string)
        except common.exceptions.ElementNotVisibleException:
            "print no key statistics"
    except (common.exceptions.NoSuchElementException):
            # print "either request rejected or company deleted "
        print "request rejected"
        sleep(180)
        driver.refresh()
        requestadjustment(i)


data1 = []

with open('ASXListedCompanies.csv') as csvDataFile:
    csvReader = csv.reader(csvDataFile)
    i = 0
    for row in csvReader:
        if(i > 2):
            data1.append(row)
        i += 1
codes = []
for i in data1:
    codes.append(i[1])

chrome_options = Options()
# chrome_options.add_argument("--headless")

driver = webdriver.Chrome(chrome_options=chrome_options)
driver.implicitly_wait(5)  # seconds
counter = 0
click = 0
asxpri = []
marketcapasx = []
dividend = []
keystatisticsasx = []
announcementsasx = []

for i in codes:
    print counter
    driver.get("https://www.asx.com.au/asx/share-price-research/company/%s" % i)
    try:
        try:
            checkincon = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/h2').text
            print "company not found"
        except common.exceptions.NoSuchElementException:
            print "company found"
            checkincon = "found"
        driver.find_element_by_xpath(
            '//*[@id="anaylst-research"]/company-research/div[2]/ul/li/a/span[1]').click()
        mainwindow = driver.window_handles[0]
        popupwindow = driver.window_handles[1]
        driver.switch_to.window(popupwindow)
        if (click == 0 and checkincon == "found"):
            driver.find_element_by_xpath(
                '/html/body/div/form/input[2]').click()
            click = 1
        pdfurl = driver.current_url
        with open('asxannreport.csv', mode='a') as annual:
            writer = csv.writer(annual, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, pdfurl
            writer.writerow(string)
        driver.close()
        driver.switch_to.window(mainwindow)
    except common.exceptions.NoSuchElementException:
        pdfurl = "-"
        print "no annual pdf"
    # annual report

    print "Fetching %s data \n ------------------------------------" % i
    # price
    try:
        price = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[1]/td[1]/span').text
        with open('asxprices.csv', mode='a') as doc:
            writer = csv.writer(doc, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, price
            writer.writerow(string)
        # marketcap
        marketcap = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[2]/td[1]/div/span').text
        with open('asxmarketcaps.csv', mode='a') as doc:
            writer = csv.writer(doc, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, marketcap
            writer.writerow(string)
        # dividend
        mostrecent = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[2]/span[1]').text
        exdate = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]').text
        paydate = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[3]/td[2]').text
        franking = driver.find_element_by_xpath(
            '//*[@id="information-column"]/div[4]/div[1]/div[1]/company-summary/table/tbody/tr[3]/td[2]/table/tbody/tr[4]/td[2]').text
        print "price :", price, '\n', "marketcap :", marketcap, '\n', "pdfurl:", pdfurl, '\n'",mostrecent :", mostrecent, '\n', "exdate :", exdate, '\n', "paydate :", paydate, '\n', "franking :", franking
        with open('asxdividens.csv', mode='a') as doc:
            writer = csv.writer(doc, delimiter=",",
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
            string = i, mostrecent, exdate, paydate, franking
            writer.writerow(string)

        # key statistics
        try:
            driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[3]/ul/li[2]/a').click()
            driver.implicitly_wait(5)
            # #day
            openprice = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[3]/td[2]/span').text
            highprice = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[4]/td[2]/span').text
            lowprice = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[5]/td[2]/span').text
            volume = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[6]/td[2]/span').text
            bid = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[7]/td[2]/span').text
            offer = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[8]/td[2]/span').text
            numerofshares = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[9]/td[2]/span').text
            # year
            prevclose = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[3]/td[4]/span').text
            week52h = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[4]/td[4]/span').text
            week52l = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[5]/td[4]/span').text
            averagevol = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[6]/td[4]/span').text
            # ratios
            pore = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[3]/td[6]/span').text
            eps = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[4]/td[6]/span').text
            ady = driver.find_element_by_xpath(
                '//*[@id="information-column"]/div[4]/div/div[3]/table/tbody/tr[5]/td[6]/span').text
            keystatisticsasx.append(i+","+openprice+","+highprice+","+lowprice+","+volume+","+bid+","+offer +
                                    ","+numerofshares+","+prevclose+","+week52h+","+week52l+","+averagevol+","+pore+","+eps+","+ady)
            print openprice, "\n", highprice, "\n", lowprice, "\n", volume, "\n", bid, "\n", offer, "\n", numerofshares, "\n", prevclose, "\n", week52h, "\n", week52l, "\n", averagevol, "\n", pore, "\n", eps, "\n", ady
            with open('asxkeystatistics.csv', mode='a') as doc:
                writer = csv.writer(doc, delimiter=",",
                                    quotechar='"', quoting=csv.QUOTE_MINIMAL)
                string = i, openprice, highprice, lowprice, volume, bid, offer, numerofshares, prevclose, week52h, week52l, averagevol, pore, eps, ady
                writer.writerow(string)
        except common.exceptions.ElementNotVisibleException:
            "print no key statistics"
    except (common.exceptions.NoSuchElementException):
        print 'click :', click
        if checkincon == "Sorry we cannot find that company.":
            print "company not found"
        else:
            print "request rejected"
            sleep(180)
            driver.refresh()
            requestadjustment(i)
    # announcements table
    driver.get(
        "https://www.asx.com.au/asx/statistics/announcements.do?by=asxCode&asxCode=%s&timeframe=D&period=M6" % i)

    mainwindow = driver.window_handles[0]
    announcements = []
    try:
        table = driver.find_element_by_xpath(
            '//*[@id="content"]/div/announcement_data/table/tbody')
        a = 1
        for tr in table.find_elements_by_tag_name("tr"):
            string = tr.text
            string = string.splitlines()
            announcement = []
            anndate = string[0]
            check = string[1].split("   ")
            if len(check) == 2:
                anntitle = check[1]
                anntime = check[0]
            else:
                anntime = string[1]
                anntitle = string[2]
            announcement.append(i)
            announcement.append(anndate)
            announcement.append(anntime)
            announcement.append(anntitle)

            driver.find_element_by_xpath(
                '//*[@id="content"]/div/announcement_data/table/tbody/tr[%s]/td[3]/a' % a).click()
            popupwindow = driver.window_handles[1]
            driver.switch_to.window(popupwindow)
            if click == 0:
                driver.find_element_by_xpath(
                    '/html/body/div/form/input[2]').click()
                click += 1
            annpdfurl = driver.current_url
            announcement.append(annpdfurl)
            announcements.append(announcement)
            with open('asxannouncements.csv', mode='a') as asxprices:
                writer = csv.writer(asxprices, delimiter=",",
                                    quotechar='"', quoting=csv.QUOTE_MINIMAL)
                string = i, anndate, anntime, anntitle, annpdfurl
                writer.writerow(string)
            a = a+1
            driver.close()
            driver.switch_to.window(mainwindow)
    except common.exceptions.NoSuchElementException:
        try:
            checkin = driver.find_element_by_xpath(
                '/html/body/div/div/h1').text
            print "request", checkin
            sleep(30)
            requestadjustment(i, click)
        except common.exceptions.NoSuchElementException:
            print "no announcments"
    counter += 1
driver.close()
