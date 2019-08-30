from html.parser import HTMLParser
import urllib
import urllib.request
from urllib.request import urlopen  
from urllib import parse
from bs4 import BeautifulSoup
from pandas import DataFrame
import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta, date
from time import localtime, strftime
import time
import sys
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import csv
import os.path

MARRIOTT_HOTELS = {
    "112809", #Courtyard by Marriott Seattle Southcenter
    "153220", #Courtyard by Marriott Sea-Tac
    "112831", #Residence inn Seattle South
}

HILTON_HOTELS = {
    "162077", #embassy suites by Hilton
    "224388", #doubletree suites by Hilton Seattle Airport
    "129030", #homewood suites by Hilton
    "478013", #home2 suites by Hilton
    "139252", #Hilton garden inn renton
    "124575", #hampton inn Seattle/South Center
    "408064", #hampton inn & suites Seattle-Airport
    "107289", #hampton inn seattle/airport
}

CHOICE_HOTELS = {
    "161820", #sleep inn seatac
    "205010", #comfort suites airport
}

IHG_HOTELS = {
    "693766", #holiday inn express Seattle South
}

WYNDHAM_HOTELS = {
    "208159", #ramada by Wyndham Tukwila Southcenter    
}

INTERURBAN_HOTELS = {
    "712914880", #hotel interurban
}

WOODSPRING_HOTELS = {
    "753362272", #woodspring suites
}

TRACKED_HOTELS = { 
    "712914880": "Hotel Interurban",
    "753362272": "WoodSpring Suites Seattle Tukwila",
    "112809": "Courtyard by Marriott Seattle Southcenter",
    "153220": "Courtyard by Marriott Seattle Sea-Tac Area",
    "112831": "Residence Inn Seattle South/Tukwila",
    "162077": "Embassy Suites by Hilton Seattle Tacoma International Airport",
    "224388": "DoubleTree Suites by Hilton Seattle Airport - Southcenter",
    "129030": "Homewood Suites by Hilton Seattle-Tacoma Airport/Tukwila",
    "478013": "Home2 Suites by Hilton Seattle Airport",
    "139252": "Hilton Garden Inn Seattle/Renton",
    "124575": "Hampton Inn Seattle/Southcenter",
    "408064": "Hampton Inn & Suites Seattle-Airport/28th Ave",
    "107289": "Hampton Inn Seattle/Airport",
    "161820": "Sleep Inn - SeaTac Airport",
    "205010": "Comfort Suites Airport",
    "115887": "Comfort Inn & Suites Sea-Tac Airport",
    "693766": "Holiday Inn Express & Suites Seattle South - Tukwila",
    "208159": "Ramada by Wyndham Tukwila Southcenter",
    "141915": "Ramada by Wyndham SeaTac Airport North",
}

MARRIOTT_TO_HOTELS = {
    "SEASC": "112809",
    "SEAWV": "153220",
    "SEASO": "112831",
}

HILTON_TO_HOTELS = {
    "SEATUES": "162077", #embassy suites by Hilton
    "SEASPDT": "224388", #doubletree suites by Hilton Seattle Airport
    "SEATKHW": "129030", #homewood suites by Hilton
    "SEAUKHT": "478013", #home2 suites by Hilton
    "SEARHGI": "139252", #Hilton garden inn renton
    "SEASOHX": "124575", #hampton inn Seattle/South Center
    "SEAIAHX": "408064", #hampton inn & suites Seattle-Airport
    "SEASTHX": "107289", #hampton inn seattle/airport
}

soldOutHotels = set()

class Hotel:
    id = ""
    name = ""
    price = ""
    address = ""
    sitePrice = ""

    def __init__(self, name):
        self.name = name

    def __str__(self):
        price = self.price
        if price == "":
            price = "S/O"

        sitePrice = self.sitePrice
        if sitePrice == "":
            sitePrice = "S/O"

        return "============{0}============\naddress:{1}\n3rd party price:{2}\nsite price:{3}\n\n".format(
            self.name,
            self.address,
            price,
            sitePrice)

def scrapeHotelsCom(checkin, checkout, driver):
    checkinStr = checkin.strftime("%Y-%m-%d")
    checkoutStr = checkout.strftime("%Y-%m-%d")
    url = "https://www.hotels.com/search.do?resolved-location=CITY%3A1467818%3AUNKNOWN%3AUNKNOWN&destination-id=1467818&q-destination=Tukwila,%20Washington,%20United%20States%20of%20America&q-rooms=1&q-room-0-adults=1&q-room-0-children=0&q-check-in=" + checkinStr + "&q-check-out=" + checkoutStr
    driver.get(url)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight/4);")
    time.sleep(1)
    for i in range(10):
        # Scroll down to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # Wait to load page
        time.sleep(1)
    content = driver.page_source
    soup = BeautifulSoup(content,features="lxml")
    hotelDict = {}
    output.write("check-in:" + checkinStr + "\n")
    output.write("check-out:" + checkoutStr + "\n")
    for hotelObj in soup.findAll("li", "hotel"):
        hotelId = hotelObj.get("data-hotel-id")
        if hotelId not in TRACKED_HOTELS:
            continue
        hotelName = hotelObj.get("data-title")
        if not hotelName:
            hotelName = hotelObj.find("a", "property-name-link").text
        hotel = Hotel(hotelName)
        hotel.address = hotelObj.find("span", "address").text
        if hotelObj.find("p", "sold-out-message"):
            hotel.price = ""
        else:
            hotel.price = (hotelObj.find("div", "price").find("ins") or hotelObj.find("div", "price").find("strong")).text
        hotel.id = hotelId
        hotelDict[hotelId] = hotel
    return hotelDict

def scrapeMarriottCom(checkin, checkout, driver):
    checkinStr = checkin.strftime("%m-%d")
    checkoutStr = checkout.strftime("%m-%d")
    url = "https://www.marriott.com/search/findHotels.mi"
    driver.get(url)
    checkin = driver.find_elements_by_class_name("ccheckin");
    if len(checkin) > 0:
        element = checkin[0]
        element.click()
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        element.send_keys(checkinStr)
        element.send_keys(Keys.TAB)
    checkout = driver.find_elements_by_class_name("ccheckout");
    if len(checkout) > 0:
        element = checkout[0]
        element.click()
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        element.send_keys(checkoutStr)
        element.send_keys(Keys.TAB)
    destination = driver.find_elements_by_xpath("//input[@data-target='destination']")
    if len(destination) > 0:
        element = destination[0]
        element.click()
        element.send_keys("Tukwila, WA")
        element.send_keys(Keys.ENTER)
    time.sleep(2)
    records = driver.find_elements_by_class_name("property-records")
    while(len(records) < 1):
        time.sleep(2)
        records = driver.find_elements_by_class_name("property-records")
    content = driver.page_source
    soup = BeautifulSoup(content, features="lxml")
    for hotelObj in soup.findAll("div", {"class":"property-record-item"}):
        marriottId = hotelObj.get("data-marsha")
        hotelId = MARRIOTT_TO_HOTELS.get(marriottId, "")
        if hotelId == "":
            continue
        price = hotelObj.find("span", "t-price").text
        hotelDict[hotelId].sitePrice = "$" + price
    return

def scrapeHiltonCom(checkin, checkout, driver):
    checkinStr = checkin.strftime("%m-%d-%Y")
    checkoutStr = checkout.strftime("%m-%d-%Y")
    url = "https://hiltonhonors3.hilton.com/en_US/hh/search/findhotels/index.htm"
    driver.get(url)
    destination = driver.find_element_by_id("hotelSearchOneBox");
    destination.click()
    destination.send_keys("Tukwila, WA")
    driver.find_element_by_xpath("//select[@name='radiusFromLocation']/option[@value='5']").click()
    checkin = driver.find_element_by_id("checkin")
    checkin.click()
    checkin.clear()
    checkin.send_keys(checkinStr)
    checkout = driver.find_element_by_id("checkout")
    checkout.click()
    checkout.clear()
    checkout.send_keys(checkoutStr)
    driver.find_element_by_xpath("//select[@id='room1Adults']/option[@value='1']").click()
    driver.find_element_by_xpath("//a[@class='linkBtn']").click()
    time.sleep(2)
    records = driver.find_elements_by_class_name("sResults")
    while(len(records) < 1):
        time.sleep(2)
        records = driver.find_elements_by_class_name("sResults")
    content = driver.page_source
    soup = BeautifulSoup(content, features="lxml")
    for hotelObj in soup.findAll("div", {"class":"sResult"}):
        hotelSelector = hotelObj.find("div", id=lambda x: x and x.startswith("quickLook"))
        if hotelSelector != "None":
            hiltonId = hotelSelector["id"][9:]
        else:
            continue
        hotelId = HILTON_TO_HOTELS.get(hiltonId, "")
        if hotelId == "":
            continue
        hotelDict[hotelId].sitePrice = hotelObj.find("h3", {"class": "statusPrice"}).find("ins").text
    return

# Set the check in and check out dates
today = date.today()
currentTime=strftime("%H%M", localtime())
try:
    daysOut = int(input("How many days out do you want to check? "))
except ValueError:
    daysOut = 3
checkin = (today + timedelta(days=daysOut))
checkout = (today + timedelta(days=daysOut+1))
checkinStr = checkin.strftime("%m-%d-%Y")
checkoutStr = checkout.strftime("%m-%d-%Y")

# Set up WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--ignore-certificate-errors")
options.add_argument("--test-type")

# We have to define a window size so that headless browser can scroll
options.add_argument("window-size=1920,1080")
options.add_argument("no-cors")

# Add headless browser
#options.add_argument("headless")
driver = webdriver.Chrome('selenium/webdriver/chromedriver.exe', options=options)
output  = open("output.txt", "w")
hotelDict = scrapeHotelsCom(checkin, checkout, driver)
for id, name in TRACKED_HOTELS.items():
    if id not in hotelDict:
        hotelDict[id] = Hotel(name)
scrapeMarriottCom(checkin, checkout, driver)
scrapeHiltonCom(checkin, checkout, driver)
names = []
addresses = []
hotelPrices = []
brandPrices = []
for hotel in hotelDict.values():
    names.append(hotel.name)
    addresses.append(hotel.address)
    hotelPrices.append(hotel.price)
    brandPrices.append(hotel.sitePrice)
    output.write(hotel.__str__())
df = DataFrame({ 'Checkin': checkinStr, 'Checkout': checkoutStr, 'Hotel Name': names, 'Address': addresses, '3rd Party': hotelPrices, 'Site': brandPrices })
filename = today.strftime("%m%d%Y") + ".xlsx"
if os.path.isfile(filename):
    book = load_workbook(filename)
    writer = pd.ExcelWriter(filename, engine = 'openpyxl')
    writer.book = book
    df.to_excel(writer, sheet_name=currentTime, index=False)
    writer.save()
    writer.close()
else:
    df.to_excel(filename, sheet_name=currentTime, index=False)
driver.close()