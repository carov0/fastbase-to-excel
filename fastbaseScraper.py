import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

# Prompts the user to enter the URL of the website and fetches the webpage content
URL = input("Enter the URL of the website.")
page = requests.get(URL)

# Creates a BeautifulSoup object to parse the HTML content of the webpage 
soup = BeautifulSoup(page.content, "html.parser")

# Finds all elements with the class representing each category
companies = soup.find_all("div", class_="c_name cmp_clmn_dots")
streets = soup.find_all("td", class_="pad_lt td_cinx_street")
cities = soup.find_all("td", class_="pad_lt td_cinx_city")
websites = soup.find_all("td", class_="pad_lt td_cinx_web")
phones = soup.find_all("td", class_="pad_lt td_cinx_phone")
names = soup.find_all("td", class_="pad_lt td_cinx_con_per")
emails = soup.find_all("td", class_="pad_lt td_cinx_email")

# Creates lists to later store each element in every category
companyList = []
streetList = []
cityList = []
websiteList = []
phoneList = []
nameList = []
emailList = []

# Extracts the relevant data from the parsed HTML elements and adds them to the previous lists 
for company in companies:
    companyList.append(company.string)
for street in streets:
    streetList.append(street.string)
for city in cities:
    cityList.append(city.string)
for website in websites:
    websiteList.append(website.string)
for phone in phones:
    phoneList.append(phone.string)
for name in names:
    nameList.append(name.string)
for email in emails:
    emailList.append(email.string)

# Orders the information in a dictionary which will be used to export the data to Excel
data = {
    'COMPANY': companyList,
    'ADDRESS': streetList,
    'CITY': cityList,
    'WEBSITE': websiteList,
    'PHONE': phoneList,
    'FULL NAME': nameList,
    'EMAIL': emailList,
    'URL': URL,
}

# Creates a dataframe from the data dictionary and exporting it to an Excel file
dataframe = pd.DataFrame(data)
documentName = input("Enter a title for your document.")
dataframe.to_excel(f"{documentName}.xlsx", index=False)

# Result message
print("Your document is ready! You will find it in the same folder where this program is stored.")


# Test URL "https://www.fastbase.com/countryindex/Finland/L/Laser-tag-center"
