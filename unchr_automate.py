from selenium import webdriver
import json
import os
import time
from pdf2docx import Converter
import pandas as pd
import numpy as np
from pandas.io.json import json_normalize
import json
from urllib.request import urlopen
from docx import Document
import pycountry
import tkinter as tk
from tkinter import filedialog

# Create a Tkinter window
root = tk.Tk()

# Prompt the user to select a folder
destination_path = filedialog.askdirectory(initialdir="C:/",
                                      title="Choose a folder",
                                      mustexist=True)

# put the today's date in the file name
today_date = time.strftime("%d%m%Y-%H%M%S")
filename = "ukr_data" + "_" + today_date + ".pdf"

options = webdriver.ChromeOptions()
settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
prefs = {'savefile.default_directory': destination_path, 'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
options.add_experimental_option('prefs', prefs)
options.add_argument('--kiosk-printing')
driver_path = r'C:\Users\Esprimo Q920\Downloads\chromedriver_win32\chromedriver.exe'
driver = webdriver.Chrome(options=options, executable_path=driver_path)
driver.get("https://app.powerbi.com/view?r=eyJrIjoiZDBlM2EwOWMtMDk2Mi00ZDc4LTliYWUtZTNjMmNlN2ZmY2Y4IiwidCI6ImU1YzM3OTgxLTY2NjQtNDEzNC04YTBjLTY1NDNkMmFmODBiZSIsImMiOjh9")
time.sleep(5)
driver.execute_script("document.title = \'{}\'".format(filename))
time.sleep(2)
driver.execute_script('window.print();')

time.sleep(2)
driver.quit()
print('PDF successful')

source_path = os.path.join(destination_path, filename)

pdf_file = source_path
output_name = "ukr_data" + "_" + today_date + ".docx"
docx_file = os.path.join(destination_path, output_name)
cv = Converter(pdf_file)
cv.convert(docx_file)      # all pages by default
cv.close()

time.sleep(2)

print('docx succesfull')

# Open the Word document
document = Document(docx_file)

# Create an empty list to store the dataframes
dataframes = []


# Loop through all the tables in the document
for table in document.tables:
  # Create an empty list to store the table data
  data = []

  # Loop through all the rows in the table
  for row in table.rows:
    # Create an empty list to store the row data
    row_data = []

    # Loop through all the cells in the row
    for cell in row.cells:
      # Append the cell text to the row data
      row_data.append(cell.text)

    # Append the row data to the table data
    data.append(row_data)

  # Create a pandas dataframe from the table data
  df = pd.DataFrame(data)

  # Append the dataframe to the list of dataframes
  dataframes.append(df)

# Assign the dataframes to variables df1, df2, and df3

df1 = dataframes[0]
df2 = dataframes[1]
df3 = dataframes[2]
df4 = dataframes[3]

df2 = df2.iloc[3:10, 0:3].reset_index(drop=True)
df2.columns = ['Country', 'Date', 'UkrainianTP']


df4.iloc[2:36, [0,1,9]].reset_index(drop=True)
df4 = df4.iloc[2:37, [0,1,9]].reset_index(drop=True)
df4.columns = ['Country', 'Date', 'UkrainianTP']


df4 = pd.concat([df2, df4], ignore_index=True)
df4 = df4.replace({"Türkiye":"Turkey", "Serbia and Kosovo: S/RES/1244 (1999)":"Kosovo"}, regex=False)
df4 = df4.drop(df4[df4['UkrainianTP']=='Not applicable'].index).reset_index(drop=True)
df4 = df4[df4["Country"].str.contains("Georgia|Turkey|United Kingdom|Kosovo|Montenegro|Azerbaijan|Bosnia and Herzegovina|Armenia|Albania")==False]


input_countries = df4['Country']

countries = {}
for country in pycountry.countries:
    countries[country.name] = country.alpha_2

codes = [countries.get(country, 'Unknown code') for country in input_countries]



df4['code_country'] = codes

df4 = df4.replace({'Unknown code':'CZ'}, regex=False)
df4['code_country'] = df4['code_country'].fillna('EL')

url = 'http://ec.europa.eu/eurostat/wdds/rest/data/v2.1/json/en/tps00001?filterNonGeo=1&precision=1&lastTimePeriod=1&shortLabel=1'
# read the JSON file into a Python object
with urlopen(url) as response:
    source = response.read()
    
d = json.loads(source)

a = d['dimension']['geo']['category']['index']
pop_code = pd.DataFrame(list(a.items()),columns = ['code',"index"]) 
pop_value = pd.DataFrame(list(d['value'].items()),columns = ['index',"Population"]) 
pop_value['index'] = pop_value['index'].astype('int64')

pop = pop_value.merge(pop_code, how='left', on='index').reset_index(drop=True).drop('index', axis=1)
pop.replace('EL', 'GR', inplace=True)

df4 = df4.merge(pop, how='left', left_on='code_country', right_on='code').drop(['code_country', 'code'], axis=1)
df4['UkrainianTP']=df4['UkrainianTP'].replace(',', '', regex=True).astype('float64')
df4['UKRper100'] = (df4['UkrainianTP']/df4['Population'])*100

df5 = df4.drop('Date', axis=1)
df5 =  df5.rename(columns={'UkrainianTP':'Ukrainians under TP', 'UKRper100':'In per cent'})
df5.head()

df5['In per cent'] = df5['In per cent'].round(decimals=2)
df5['Population'] = df5['Population']/1000000
df5 =  df5.rename(columns={'Population':'Population (mio)'})
df5['Population (mio)'] = df5['Population (mio)'].round(decimals=2)
df5 = df5.sort_values(by=['In per cent','Ukrainians under TP', 'Population (mio)'], ascending=False, ignore_index=True)
df5=df5.reindex(columns=['Country', 'Population (mio)', 'Ukrainians under TP', 'In per cent'])
group1 = df5[df5['In per cent']<1]
group2 = df5[(df5['In per cent']>=1) & (df5['In per cent']<2)]
group3 = df5[df5['In per cent']>=2]

excel_output_name = "ukr_data" + "_" + today_date + ".xlsx"
excel_file = os.path.join(destination_path, excel_output_name)

ukr = pd.DataFrame()
ukr.to_excel(excel_file)

# create a excel writer object
with pd.ExcelWriter(excel_file) as writer:
   
    # use to_excel function and specify the sheet_name and index
    # to store the dataframe in specified sheet
    group1.to_excel(writer, sheet_name="<1", index=False)
    group2.to_excel(writer, sheet_name="1>=x<2", index=False)
    group3.to_excel(writer, sheet_name="2>=x", index=False)
    df5.to_excel(writer, sheet_name="All", index=False)

print('excel successful')