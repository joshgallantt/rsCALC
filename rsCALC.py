import os
import time
import csv
import sys
import re
import sys
import pathlib
import calendar
import pandas as pd
import shutil
import platform
import tkinter as tk
from tabulate import tabulate
from PIL import Image, ImageTk
from tkinter.filedialog import askopenfile
from tkinter import ttk
from tkinter import * 
from tkcalendar import Calendar, DateEntry
from collections import Counter
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from pathlib import Path


def clear():
    return os.system('cls' if os.name == 'nt' else 'clear')

def wait(seconds):
    return time.sleep(seconds)

def download_wait(directory = str(os.getcwd())):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(directory):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

def login():
    driver.get("https://auth.rewardstyle.com/login/")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".mt-5 > .btn__content")))
    driver.find_element(By.NAME, "username").send_keys(username)
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".mt-5 > .btn__content").click()
    driver.get("https://www.rewardstyle.com/affiliate-rewards")


def read_from_calendar():
    wait(0.1)
    global day_css_offset
    global from_calendar_year
    global from_calendar_month
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".\\_month")))
    from_calendar_year = (driver.find_element(By.CSS_SELECTOR, ".\\_month").text[6:10])
    from_calendar_month = (driver.find_element(By.CSS_SELECTOR, ".\\_month").text[2:5])
    day_css_offset = 8


def read_to_calendar():
    wait(0.1)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".\\_month")))
    global day_css_offset
    global to_calendar_year
    global to_calendar_month
    to_calendar_year = (driver.find_element(By.CSS_SELECTOR, ".\\_month").text[6:10])
    to_calendar_month = (driver.find_element(By.CSS_SELECTOR, ".\\_month").text[2:5])
    day_css_offset = 8


def open_from_calendar():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rewards-datepicker-input-start")))
    return driver.find_element(By.CSS_SELECTOR, ".rewards-datepicker-input-start").click()


def open_to_calendar():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rewards-datepicker-input-end")))
    return driver.find_element(By.CSS_SELECTOR, ".rewards-datepicker-input-end").click()


def click_day():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".\\_day:nth-child(8)")))
    return driver.find_element(By.CSS_SELECTOR, ".\\_day:nth-child("+str(day_css_offset)+")").click()


def selected_day_on_widget():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".\\_day:nth-child(8)")))
    return int(driver.find_element(By.CSS_SELECTOR, ".\\_day:nth-child("+str(day_css_offset)+")").text)


def user_start_date_month():
    return str(users_start_date.strftime("%b")).upper()


def user_end_date_month():
    return str(users_end_date.strftime("%b")).upper()


def click_previous():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".\\_previous")))
    return driver.find_element(By.CSS_SELECTOR, ".\\_previous").click()


def click_export():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".date-picker-submit")))
    driver.find_element(By.CSS_SELECTOR, ".date-picker-submit").click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".date-picker-submit")))
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "export-button")))
    wait(3)
    driver.find_element(By.ID, "export-button").click()
    download_wait()

def find_last_day_of_month(to_calendar_month):
    calendar_dict = {
    'JAN':1, 'FEB':2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
    'JUL': 7, 'AUG': 8,'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
    }
    month_num = calendar_dict[to_calendar_month] 
    return calendar.monthrange(int(to_calendar_year), month_num)[1]


def delete_all_html_in_cwd():
    for file in os.listdir(os.getcwd()):
        if file.endswith(".html"):
            os.remove(os.path.join(file))

def delete_all_csv_in_cwd():
    for file in os.listdir(os.getcwd()):
        if file.endswith(".csv"):
            os.remove(os.path.join(file))


def convert_xls_to_html_in_cwd():
    for file in os.listdir(os.getcwd()):
        os.path.splitext(file)
        os.rename(file, file.replace('.xls', '.html'))


def soup_to_csv(outputFilename, soupFunction):

    soup_export_no_empty_lists = [lists for lists in soupFunction if lists != []]

    with open(outputFilename, 'w', encoding ='utf-8') as soup_export:
        writer = csv.writer(soup_export)
        writer.writerows(soup_export_no_empty_lists)
        pass

    if outputFilename == 'rates.csv':
        # pass
        with open(outputFilename, "r", encoding ='utf-8') as text:
            text = ''.join([i for i in text]).replace(',"',',')
            text = ''.join([i for i in text]).replace(']"\n\n','\n')
            text = ''.join([i for i in text]).replace(",['", ',')
            text = ''.join([i for i in text]).replace("%'", '')
            text = ''.join([i for i in text]).replace('<td>', '')
            text = ''.join([i for i in text]).replace('</td>', '')
            text = ''.join([i for i in text]).replace('&amp;', '&')
            pass

        with open(outputFilename,"w") as edited:
            edited.writelines(text)
            pass
    else:
        with open(outputFilename, "r", encoding ='utf-8') as text1:
            text1 = ''.join([i for i in text1]).replace('Â£','')
            text1 = ''.join([i for i in text1]).replace('"\n','')
            text1 = ''.join([i for i in text1]).replace('\n",',',')
            text1 = ''.join([i for i in text1]).replace('\n\n','\n')
            pass

        with open(outputFilename,"w") as edited1:
                edited1.writelines(text1)
                pass


def get_soup(data_or_rates):

    if data_or_rates == 'rates.csv':
        output_rows = []
        for filename in os.listdir(os.getcwd()):
            if filename.endswith("age.html"):

                html = open(filename).read()
                soup = BeautifulSoup(html, features= "lxml")
                table = soup.find("table")

                for table_row in table.findAll('tr'):
                    columns = table_row.findAll('td')
                    output_row = []
                    for column in columns:
                        columns[1] = [w.replace(' - ', ',') for w in columns[1]]
                        output_row.append(column)
                    output_rows.append(output_row)
        return output_rows

    else:
        output_rows1 = []
        for filename1 in os.listdir(os.getcwd()):
            if filename1.endswith(").html"):

                html1 = open(filename1).read()
                soup1 = BeautifulSoup(html1, features= "lxml")
                table1 = soup1.find("table")

                for table_row1 in table1.findAll('tr'):
                    columns1 = table_row1.findAll('td')
                    output_row1 = []
                    for column1 in columns1:
                        output_row1.append(column1.text)
                    output_rows1.append(output_row1)
        return output_rows1


def download_data():

    global day_css_offset
    global from_calendar_year
    global from_calendar_month
    global to_calendar_year
    global to_calendar_month
    global progress

    # open and read the end calendar widget for the first time, every time we read them we set the div offest to 8
    open_to_calendar()
    read_to_calendar()

    # while either the widgets month or year dont match our END date, we move it back until it does.
    while int(to_calendar_year) != users_end_date.year or to_calendar_month != user_end_date_month():
        click_previous()
        read_to_calendar()

    # then we find their end date and click it
    while selected_day_on_widget() != 1:
        day_css_offset += 1

    while selected_day_on_widget() < users_end_date.day:
        day_css_offset += 1

    click_day()


    #loop to iterate through the months, with break if it's the last month
    while True:

        #open the start widget and read it and set the first date to div offset 8.
        wait(0.1)
        open_from_calendar()
        wait(0.1)
        read_from_calendar()
        wait(0.1)

        #while widgets month or year dont match our START date, we will have to download from the first.
        progress['value'] += 10
        progress.update()

        while int(from_calendar_year) != users_start_date.year or from_calendar_month != user_start_date_month():

            #select the first of the month
            while selected_day_on_widget() != 1:
                day_css_offset += 1
            click_day()

            #if the end widget month/year dont match our users exactly, we need to select end of month:
            open_to_calendar()
            read_to_calendar()

            if int(to_calendar_year) != users_end_date.year and to_calendar_month != user_end_date_month():

                while selected_day_on_widget() != 1:
                    day_css_offset += 1

                while selected_day_on_widget() < find_last_day_of_month(to_calendar_month):
                    day_css_offset += 1
                click_day()

            #export the month and set end widget to the last day of the previous month
            click_export()

            open_to_calendar()
            click_previous()
            read_to_calendar()

            while selected_day_on_widget() != 1:
                day_css_offset += 1

            while selected_day_on_widget() < find_last_day_of_month(to_calendar_month):
                day_css_offset += 1
            click_day()
            break

        #check to see if this month is last month to download, if so, set the user start date and download 
        open_from_calendar()
        read_from_calendar()

        if from_calendar_month == user_start_date_month() and int(from_calendar_year) == users_start_date.year:

            while selected_day_on_widget() != 1:
                day_css_offset += 1

            while selected_day_on_widget() < users_start_date.day:
                day_css_offset += 1
            click_day()

            click_export()

            break


def download_rates():
    driver.get("https://www.rewardstyle.com/ads/rates?s=0")
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".even-row:nth-child(2) > td:nth-child(1)")))

    with open('page.html', 'w+', encoding="utf-8") as f:
        f.write(driver.page_source)
        f.close()



def generate_report(request):

    pd.set_option('precision', 2)
    pd.options.display.float_format = '£{:,.2f}'.format
    df1 = pd.read_csv('user_data.csv', encoding='utf-8', delimiter=',', names=['Type', 'Advertiser', 'Date', 'Link', 'Influencer', 'Product', 'Status', 'Our Earnings', 'Payments'], skip_blank_lines=True, skipinitialspace=True, engine='python',header=None) 
    df2 = pd.read_csv('rates.csv', encoding='utf-8', delimiter=',', names=['Advertiser', 'Rate 1', 'Rate 2'], skip_blank_lines=True, skipinitialspace=True, engine='python',header=None) 
    
    df = pd.merge(df1, df2,  
                        on ='Advertiser',  
                        how ='inner')
    df


    df['Rate'] = ((df['Rate 1'] + df['Rate 2']) / 2)/100
    del df['Rate 1']
    del df['Rate 2']
    df['Advertisers Earnings'] = df['Our Earnings'] / df['Rate']


    a = f'Generated report between {users_start_date} and {users_end_date}: '
    if request == 'a':
        return a



    #Sum of your open earnings
    total_earnings_for_date_range = round(df['Our Earnings'].sum(),2)
    
    b = '\nTotal open earnings of {:,.2f}'.format(total_earnings_for_date_range)
    if request == 'b':
        return b

    #Sum of all brands estimated earnings
    total_advertiser_earnings_for_date_range = round(df['Advertisers Earnings'].sum(),2)
    c = '\nBrands earned an estimated {:,.2f}'.format(total_advertiser_earnings_for_date_range)
    if request == 'c':
        return c

    #Total Earnings by Advertiser
    total_earnings_by_advertiser = df.groupby(['Advertiser'])[["Our Earnings", "Advertisers Earnings"]].sum().sort_values(by='Advertisers Earnings', ascending=False).round(decimals =2)
    d = '\nTotal Earnings by Advertiser: '
    e = total_earnings_by_advertiser
    if request == 'd':
        return d
    if request == 'e':
        return tabulate(total_earnings_by_advertiser, headers=["Advertiser","Our Earnings", "Advertiser Earnings"])

    #Top 5 products in period by number sold
    f = '\nTop Products by Numbers Sold: '
    sales = df['Type']== 'Sale Commission'
    df_sales = df[sales]
    top_10_products_by_number = df_sales.groupby(['Advertiser','Product']).agg({'Product': 'count', 'Our Earnings': 'sum'}).rename(columns={'Product':'Number Sold'}).reset_index().sort_values(by='Number Sold', ascending=False).head(10)

    g = top_10_products_by_number
    if request == 'f':
        return f
    if request == 'g':
        return tabulate(top_10_products_by_number, headers=["Brand","Product", "# Sold", "Our Earnings"], showindex=False)

    #Top 5 products in period by comission
    h = '\nTop Products by Commission: '
    sales = df['Type']== 'Sale Commission'
    df_sales = df[sales]
    top_10_products_by_earned = df_sales.groupby(['Advertiser','Product']).agg({'Product': 'count', 'Our Earnings': 'sum'}).rename(columns={'Product':'Number Sold'}).reset_index().sort_values(by='Our Earnings', ascending=False).head(10)
    j = top_10_products_by_earned.to_string(index=False)
    if request == 'h':
        return h
    if request == 'j':
        return tabulate(top_10_products_by_earned, headers=["Brand","Product", "# Sold", "Our Earnings"], showindex=False)


    #Most Refunded Products:
    k = '\nMost Refunded Products: '
    sales = df['Type']== 'Sale Return'
    df_sales = df[sales]
    top_10_ref = df_sales.groupby(['Advertiser','Product']).agg({'Product': 'count', 'Our Earnings': 'sum'}).rename(columns={'Product':'Number Sold'}).reset_index().sort_values(by='Number Sold', ascending=False).head(10)
    l = top_10_ref
    if request == 'k':
        return k
    if request == 'l':
        return tabulate(top_10_ref, headers=["Brand","Product", "# Refunded", "Loss"], showindex=False)


def initdriver():

    global from_calendar_year
    global from_calendar_month
    global to_calendar_year
    global to_calendar_month
    global day_css_offset
    global options

    from_calendar_year = ''
    from_calendar_month = ''
    to_calendar_year = ''
    to_calendar_month = ''
    day_css_offset = 8
    options = Options()
    options.add_argument('--headless')
    options.add_argument('window-size=1920x1080')
    options.add_argument("--disable-notifications")
    options.add_argument('log-level=3')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option("prefs", {
    "download.default_directory": str(os.getcwd()),
    "download.prompt_for_download": False,
    "download._upgrade": True,
    "safebrowsing.enabled": True
    })

    global driver

    myos = platform.system()
    cpu = platform.architecture()

    if myos == 'Windows':
        osdependant = "drivers/win"

    if myos == 'Linux':
        osdependant = "drivers/linux"

    if myos == 'Darwin':
        osdependant = "drivers/mac64"
    
    driver = webdriver.Chrome(options=options, executable_path=osdependant)

def start():

    global root
    global username
    global password
    global users_start_date
    global users_end_date
    global progress
    global text_box
    global EventScrollBar

    # status to be implemented
    # Label(root, text='status...', font=('TkDefaultFont', 12, 'normal')).place(x=40, y=385)
    # Label(root, text='100%', font=('TkDefaultFont', 12, 'normal')).place(x=435, y=385)
    progress=ttk.Progressbar(root, orient='horizontal', length=440, mode='determinate', maximum=100, value=1)
    progress.place(x=38, y=410)

    username = userinput.get()
    password = passinput.get()

    users_start_date = fromDate.get_date()
    users_end_date = toDate.get_date()
    
    
    try:
        if text_box.winfo_exists():
            text_box.pack_forget()
            EventScrollBar.pack_forget()
    except:
        pass


    
    progress['value'] = 5
    progress.update()
    initdriver()
    login()
    progress['value'] = 20
    progress.update()
    download_data()
    download_rates()
    progress['value'] += 10
    progress.update()
    download_wait()
    driver.close()
    progress['value'] += 10
    progress.update()
    convert_xls_to_html_in_cwd()

    soup_to_csv('user_data.csv', get_soup('user_data.csv'))

    soup_to_csv('rates.csv', get_soup('rates.csv'))
    progress['value'] +=10
    delete_all_html_in_cwd()

    progress['value'] = 100
    progress.update()
    
    root.geometry('1600x500')
    text_box = tk.Text(root, font=("Courier",8), height=32, width=150, padx= 10, pady = 10)

    EventScrollBar= tk.Scrollbar(root, command=text_box.yview, orient="vertical")

    text_box.insert(END, generate_report('a') + '\n')
    text_box.insert(END, generate_report('b') + '\n')
    text_box.insert(END, generate_report('c') + '\n')
    text_box.insert(END, generate_report('d') + '\n')
    text_box.insert(END, generate_report('e') + '\n')
    text_box.insert(END, generate_report('f') + '\n')
    text_box.insert(END, generate_report('g') + '\n')
    text_box.insert(END, generate_report('h') + '\n')
    text_box.insert(END, generate_report('j') + '\n')
    text_box.insert(END, generate_report('k') + '\n')
    text_box.insert(END, generate_report('l') + '\n')

    text_box.place(x=520, y=10)
    EventScrollBar.pack(side = RIGHT, fill = Y)  
    text_box.configure(yscrollcommand=EventScrollBar.set)

    Button(root, text='Export', width = 9, font=('TkDefaultFont', 12, 'normal'), command=lambda:export()).place(x=210, y=450)




# this is the function called when the button is clicked
def export():
    path_to_export = filedialog.asksaveasfilename(
    defaultextension='.csv', filetypes=[("csv files", '*.csv')],
    initialdir= str(Path.home()),
    title="Choose export location")
    try:
        shutil.copyfile('user_ .csv',path_to_export)
    except:
        pass


if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys._MEIPASS)
elif __file__:
    application_path = os.path.dirname(__file__)


os.chdir(str(application_path))
current_directory = os.getcwd()


delete_all_csv_in_cwd()
root = Tk()


# This is the section of code which creates the main window
root.geometry('520x500')
root.title('rsCALC - Version 0.1')
root.resizable(False, False)
root.iconphoto(False, tk.PhotoImage(file='assets/dress.png'))


# First, we create a canvas to put the picture on
logo= Canvas(root, height=128, width=128)
picture_file = PhotoImage(file ='assets/dress.png')
logo.create_image(128, 0, anchor=NE, image=picture_file)
logo.place(x=195, y=40)

userinput=Entry(root, width = 25)
userinput.place(x=235, y=190)

passinput=Entry(root, show="*", width = 25)
passinput.place(x=235, y=230)

Label(root, text='Username', font=('TkDefaultFont', 12, 'normal')).place(x=135, y=190)
Label(root, text='Password', font=('TkDefaultFont', 12, 'normal')).place(x=135, y=230)

Label(root, text='From Date', font=('TkDefaultFont', 12, 'normal')).place(x=160, y=270)
fromDate= DateEntry(width=12, borderwidth=2)
fromDate.place(x=260, y=270)

Label(root, text='To Date', font=('TkDefaultFont', 12, 'normal')).place(x=160, y=310)
toDate= DateEntry(width=12, borderwidth=2)
toDate.place(x=260, y=310)

Button(root, text='Start', width = 9, font=('TkDefaultFont', 12, 'normal'), command=lambda:start()).place(x=210, y=355)


root.mainloop()

try:
    driver.close()
except:
    pass
