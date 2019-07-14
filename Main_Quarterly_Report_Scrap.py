# -*- coding: utf-8 -*-
#Created on Sat Jul  6 12:53:25 2019
#@author: ArunabhaSarkar

# Folder Structure Required
# A Masterfolder ('Quarterly_Reports' on my PC)
# 'Quarterly_Reports' folder containing folders:
# Main, Resourses & Output
# 'Main' folder has the 2 code files titles:
## 'Main_Quarterly_Report_Scrap.py' & 'Tools_Quarterly_Report_Scrap.py'
# 'Resources' folder has the required 'gecko' driver software
# 'Output' folder (empty before first run)

# The Github upload also has 2 raw upload sample outputs named:
## 'Adani Ports and Special Economic Zone Ltd. Q_Report_2019_07_08_14_25_50.py'
## 'Asian Paints Ltd. Q_Report_2019_07_08_14_30_12.csv'

# The code to process output raw data files 
# like the above mentioned above can be found on my Github:
# https://github.com/NV2017/Cleaning_raw_Quarterly_Reports_raw_money_control_scrap_data


import os # And setting proper directory
path_2=r'C:\Users\Arunabha Sarkar\Documents\QRS\Capstones\Project Zero Sigma Hedge\Web_Scrap\Moneycontrol\Quarterly_Reports\Main'
os.chdir(path_2)
import Tools_Quarterly_Report_Scrap as SCRAPER
path_1=r'C:\Users\Arunabha Sarkar\Documents\QRS\Capstones\Project Zero Sigma Hedge\Web_Scrap\Moneycontrol\Quarterly_Reports'
os.chdir(path_1)
print('The working space directory is reset to {} for initiating scraping.'.format(path_1))
path_save = path_1 + r'\Output'
path_gecko = path_1 + r'\Resources\geckodriver'

#listname = 'Nifty_50'
#listname = 'Nifty_500'
listname = 'BSE_Nifty_500'

Company_not_found = [] # The ones that are not in moneycontrol search bar
Company_found = [] # The ones that are in moneycontrol search bar
driver = SCRAPER.startFirefoxDriver(path_gecko)
if driver is not None:

    # Getting data for Nifty 500 or Nifty 50 or NSE & BSE 500 Stocks
    Ticker_Name_List, Company_Name_List = SCRAPER.readInputFile(listname)

    # Driver to moneycontrol !
    search_bar_element, search_bar_exist = SCRAPER.To_Money_Control_Website(driver)
    
    if Company_Name_List is not None and search_bar_exist is not False:
        SCRAPER.getQuaterlyReportDataForAllCompanies(driver, Company_Name_List, 
        Ticker_Name_List, Company_found, Company_not_found, listname, path_save)

    # Shutting down the FireFox Driver 
    if driver is not None:
        SCRAPER.stopFirefoxDriver(driver)
SCRAPER.stopFirefoxDriver(driver)
print('--x--x-- END OF PROGRAM --x--x--')