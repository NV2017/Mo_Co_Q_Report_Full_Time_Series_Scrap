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



import pandas as pd
from selenium import webdriver
import time
import os
import numpy as np
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import datetime as dt
Time_Multiple = 1.5

# Night path is upper !

###############################################################################
###############################################################################
################## Function for clean starting of the driver ##################
def startFirefoxDriver(geckodriver_path):
    try:
        profile = FirefoxProfile()
        profile.set_preference("browser.helperApps.neverAsk.saveToDisk", 'application/zip')
        print('Loading Firefox Driver...')
        driver = webdriver.Firefox(firefox_profile=profile,executable_path = geckodriver_path)
        time.sleep(Time_Multiple*np.random.uniform(4.5, 5.5, size = 1))
        return driver
    except Exception as e:
        print('Error encountered while starting FireFox driver. {}'.format(e))
        return None
###############################################################################
###############################################################################
        

###############################################################################
###############################################################################
############### Function for moneycontrol search bar existance ################
def Search_Bar_Check(driver):
    time.sleep(Time_Multiple*np.random.uniform(1.5, 2.5, size = 1))
    Tester = 'XYZZY'
    search_bar_xpath = '//*[@id="search_str"]'
    try:
        Tester_element = driver.find_element_by_xpath(search_bar_xpath)
        Tester = Tester_element.text
        if Tester != 'XYZZY':
            return (Tester_element,True)
    except TimeoutException:
        print("Search bar not found in moneycontrol.com")
        return (None,False)
###############################################################################
###############################################################################
        
    
###############################################################################
###############################################################################
############## Finding out which index, NIFTY50/500/BSE500/all ################

# The input is read from my Github .csv containing company names & tickers
# Necessary changes have to be made for reading from a .csv on disk

def readInputFile(listname):
    BSE_Nifty_500 = 'https://raw.githubusercontent.com/NV2017/NIFTY_BSE_500_Combined/master/NIFTY_BSE_500_20190319.csv'
    Nifty_500 = 'https://raw.githubusercontent.com/NV2017/NIFTY500/master/ind_nifty500list_20190315.csv'
    Nifty_50 = 'https://raw.githubusercontent.com/NV2017/NIFTY500/master/20190706_ind_nifty50list.csv'

    if 'bse' in listname.lower() and 'nifty' in listname.lower().replace('_',''):
        url = BSE_Nifty_500
        df = pd.read_csv(url)
        Ticker_List = df.loc[:,'TICKERS'].values.tolist()
        Company_List = df.loc[:,'Company'].values.tolist()
        
# This is the structure to do some companies manually
        
#        Ticker_List = Ticker_List[540:]
#        Company_List = Company_List[540:]
        
#        Ticker_List = ['VGUARD','VIPIND']
#        Company_List = ['VGuard Industries','VIP Industries']

    if 'BSE' not in listname and 'nifty50' in listname.lower().replace('_',''):
        url = Nifty_50
        df = pd.read_csv(url)
        Ticker_List = df.loc[:,'Symbol'].values.tolist()
        Company_List = df.loc[:,'Company Name'].values.tolist()
        
        # To start from abrupt end
#        Ticker_List = Ticker_List[45:]
#        Company_List = Company_List[45:]
        
        # To start with known companies
#        Ticker_List = ['M&M']
#        Company_List = ['Mahindra']

    if 'BSE' not in listname and 'nifty500' in listname.lower().replace('_',''):
        url = Nifty_500
        df = pd.read_csv(url)
        Ticker_List = df.loc[:,'Symbol'].values.tolist()
        Company_List = df.loc[:,'Company Name'].values.tolist()
        
    
    return(Ticker_List,Company_List)
    
###############################################################################
###############################################################################
    
    
###############################################################################
###############################################################################
################ Function for cancelling the popular ad pop up ################
def Cancel_Adv_If_Present(driver):
                      #/html/body/div[2]/div[3]/div[2]/div[2]/div/div[1]/div[2]/span[1]/div/div/div/div[1]/div[1]/a/img
                      #/html/body/div[4]/div[3]/div[2]/div[2]/div/div[1]/div[2]/span[1]/div/div/div/div[1]/div[1]/a/img
    Add_Cross_xpath = '/html/body/div[2]/div[3]/div[2]/div[2]/div/div[1]/div[2]/span[1]/div/div/div/div[1]/div[1]/a/img'
    
    try:
        Add_Cross_element = driver.find_element_by_xpath(Add_Cross_xpath)
        Add_Cross_element.click()
    except:
        pass
###############################################################################
###############################################################################
        
    
###############################################################################
###############################################################################
###################### Typing and checking in search bar ######################
def Typing_Checking_from_moneycontrol_search_bar(driver, counter, company, Ticker_Name_List, Company_found, Company_not_found):
    search_bar = driver.find_element_by_name("search_str")
    Cancel_Adv_If_Present(driver)
    time.sleep(Time_Multiple*np.random.uniform(0.5, 1, size = 1))
    search_bar.click()
    company = company.replace(' Ltd.','').replace('(','').replace(')','')
    for j in range(len(company)): #Entering one letter at a time, like human, with random time delays between letters
        time.sleep(Time_Multiple*3*np.random.uniform(0.1, 0.8, size = 1))
        search_bar.send_keys(company[j])
    
    time.sleep(Time_Multiple*np.random.uniform(1, 1.5, size = 1))  
    
    # Finding the correct search result from drop down
    # We match 'ticker' with the search drop down list
    
    Result_css_selector_start = '.jspPane > li:nth-child('
    Result_css_selector_end = ')'
    
    start_counter = 1
    end_counter = 10
    
    # Checking only first 'end_counter' results for match with ticker
    while True:
        Search_result_trial_css_selector = Result_css_selector_start + str(start_counter) + Result_css_selector_end
        Search_result_trial_css_selector_element = ' '
        try:
            Search_result_trial_css_selector_element = driver.find_element_by_css_selector(Search_result_trial_css_selector)
        except:
            pass
        
        temporary_ticker = Ticker_Name_List[counter].split(".")[0]
        
        if temporary_ticker in Search_result_trial_css_selector_element.text:
            print(company + ' records PRESENT in moneycontrol.com')
            Company_found.append(company)
            Search_result_trial_css_selector_element.click()
            return(True)
            break
        
        if start_counter == end_counter:
            print(company + ' records ABSENT in moneycontrol.com')
            Company_not_found.append(company)
            for j in range(len(company)): #Entering one letter at a time, like human, with random time delays between letters
                time.sleep(Time_Multiple*np.random.uniform(0.1, 0.8, size = 1))
                search_bar.send_keys(Keys.BACKSPACE)
            return(False)
            break
        
        start_counter = start_counter + 1
    
###############################################################################
###############################################################################
        
        
###############################################################################
###############################################################################
################# Navigating to Quarterly Reports of company ##################
def Navigate_to_Quarterly_reports(driver, company, counter, Ticker_Name_List):
    try:
                                   #slider > dt:nth-child(9) > a:nth-child(1)
        Financials_css_selector = '#slider > dt:nth-child(9) > a:nth-child(1)'
        Financials_css_selector_element = driver.find_element_by_css_selector(Financials_css_selector)
        Cancel_Adv_If_Present(driver)
        time.sleep(Time_Multiple*np.random.uniform(0.5, 1, size = 1))
        Financials_css_selector_element.click()
        time.sleep(Time_Multiple*np.random.uniform(3.5, 4.5, size = 1))
        
        driver.execute_script("window.scrollTo(0, 0.1 * document.body.scrollHeight);")
        time.sleep(Time_Multiple*np.random.uniform(2.5, 3, size = 1))
                                         #li.act:nth-child(3) > a:nth-child(1)
                                         #ul.act > li:nth-child(3) > a:nth-child(1)
        Quarterly_Reports_css_selector = 'ul.act > li:nth-child(3) > a:nth-child(1)'
        Quarterly_Reports_css_selector_element = driver.find_element_by_css_selector(Quarterly_Reports_css_selector)
        Quarterly_Reports_css_selector_element.click()
        time.sleep(Time_Multiple*np.random.uniform(3.5, 4.5, size = 1))
        Cancel_Adv_If_Present(driver)
        
        # Checking if correct page or not
                            #/html/body/div[2]/div[3]/div[2]/div[2]/div/div[1]/div[1]
        Description_xpath = '/html/body/div[2]/div[3]/div[2]/div[2]/div/div[1]/div[1]'
        Description_element = driver.find_element_by_xpath(Description_xpath)
        
        temporary_ticker = Ticker_Name_List[counter].split(".")[0].lower()
        if Description_element and temporary_ticker in Description_element.text.lower():
            return(True)
        else:
            return(False)
    except:
        return(False)
###############################################################################
###############################################################################
        
        
###############################################################################
###############################################################################
################## Find correct 'Previous' & 'Next' button ####################
def Find_Correct_Previous_Year_Button_return_xpath(driver):
    
    # There is no clicking in this function, only collecting data
    
    Next_Button_Present = False
    Previous_Button_Present = False
    Both_Next_Previous_Present = False
    No_Buttons_Present = False
                               #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b
                               #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b
    Only_Previous_xpath =      '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b'
    # The above is same as 'Only_Next_xpath'
                               
                               #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[2]/b
                               #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[2]/b
    Previous_xpath_when_both = '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[2]/b'
    
                               #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[1]/b
                               #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[1]/b
    Next_xpath_when_both =     '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[1]/b'
    
    Only_Previous_xpath_text = ' '
    Previous_xpath_when_both_xpath_text = ' '
    Next_xpath_when_both_xpath_text = ' '
    total_buttons = ' '
    
    try:
        Only_Previous_xpath_element = driver.find_element_by_xpath(Only_Previous_xpath)
        Only_Previous_xpath_text = Only_Previous_xpath_element.text
    except:
        pass
    
    try:
        Previous_xpath_when_both_xpath_element = driver.find_element_by_xpath(Previous_xpath_when_both)
        Previous_xpath_when_both_xpath_text = Previous_xpath_when_both_xpath_element.text
    except:
        pass
    
    try:
        Next_xpath_when_both_xpath_element = driver.find_element_by_xpath(Next_xpath_when_both)
        Next_xpath_when_both_xpath_text = Next_xpath_when_both_xpath_element.text
    except:
        pass
    
    total_buttons = Only_Previous_xpath_text + ' ' + Previous_xpath_when_both_xpath_text + ' ' + Next_xpath_when_both_xpath_text
    total_buttons = total_buttons.lower()
    
    if 'next' in total_buttons and 'previous' in total_buttons:
        Next_Button_Present = True
        Previous_Button_Present = True
        Both_Next_Previous_Present = True
        No_Buttons_Present = False
    
    if 'next' in total_buttons and 'previous' not in total_buttons:
        Next_Button_Present = True
        Previous_Button_Present = False
        Both_Next_Previous_Present = False
        No_Buttons_Present = False
    
    if 'next' not in total_buttons and 'previous' in total_buttons:
        Next_Button_Present = False
        Previous_Button_Present = True
        Both_Next_Previous_Present = False
        No_Buttons_Present = False
    
    if 'next' not in total_buttons and 'previous' not in total_buttons:
        Next_Button_Present = False
        Previous_Button_Present = False
        Both_Next_Previous_Present = False
        No_Buttons_Present = True
        
    return(Next_Button_Present, Previous_Button_Present, Both_Next_Previous_Present, No_Buttons_Present)
        
###############################################################################
###############################################################################
    
    
###############################################################################
###############################################################################
######################### Only read table from page ###########################
def Only_read_table_from_page(driver):
                           #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td[
                           #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td[
    Initial_column_xpath = '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td['
    Final_column_xpath = ']'
    
    Column_Number = 0
    
    while True:
        custom_xpath = Initial_column_xpath + str(Column_Number+1) + Final_column_xpath
        exist = False
        test_string = 'XYZZY'
        try:
            custom_xpath_element = driver.find_element_by_xpath(custom_xpath)
            test_string = custom_xpath_element.text
        except:
            pass
        
        if test_string != 'XYZZY':
            Column_Number = Column_Number + 1
            exist = True
            
        if exist == False:
            break
    
    custom_columns = []
    for k in range(2,Column_Number+1):
        custom_xpath = Initial_column_xpath + str(k) + Final_column_xpath
#        print(custom_xpath)
        temp_element_text = 'XYZZY'
        try:
            temp_element = driver.find_element_by_xpath(custom_xpath)
            temp_element_text = temp_element.text
        except:
            pass
        
        if temp_element_text != 'XYZZY':
            custom_columns.append(temp_element_text)
        
#    custom_columns = custom_columns[::-1]
    custom_columns = ['Category'] + custom_columns
    
    Row_Range = list(range(0,101))
    Row_Number = 101
    
    Page_data_Output = pd.DataFrame(index = Row_Range, columns=custom_columns)
    
                         #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[
                         #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[
    Initial_cell_xpath = '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr['
    Mid_cell_xpath = ']/td['
    Final_cell_xpath = ']'
        
    for c in range(Column_Number):
        for r in range(4,Row_Number):
            cell_xpath = Initial_cell_xpath + str(r) + Mid_cell_xpath + str(c+1) + Final_cell_xpath
            text = 'XYZZY'
            try:
                text = driver.find_element_by_xpath(cell_xpath).text
            except:
                pass
            
            if text != 'XYZZY':
                Page_data_Output.iloc[r-4,c] = text
    
    
    return(Page_data_Output)

###############################################################################
###############################################################################
    
    
###############################################################################
###############################################################################
########################### Correct_Previous_Button ###########################
def Correct_Previous_Button_Click(driver, Next_Button_Present, Previous_Button_Present, Both_Next_Previous_Present):
    Cancel_Adv_If_Present(driver)
    if Next_Button_Present == False and Previous_Button_Present == True:
                             #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b
                             #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b
        correct_prev_xpath = '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b'
        correct_prev_element = driver.find_element_by_xpath(correct_prev_xpath)
        correct_prev_element.click()
    
    if Next_Button_Present == True and Previous_Button_Present == True:
                             #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[2]/b
                             #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[2]/b
        correct_prev_xpath = '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a[2]/b'
        correct_prev_element = driver.find_element_by_xpath(correct_prev_xpath)
        correct_prev_element.click()
    time.sleep(Time_Multiple*np.random.uniform(3.5, 4.5, size = 1))
    
###############################################################################
############################################################################### 
    
    
###############################################################################
###############################################################################
################### Record & Save company Quarterly Record ####################
def Record_Save_company_Quarterly_Record(driver, company, path_save):
    
    Next_Button_Present, Previous_Button_Present, Both_Next_Previous_Present, No_Buttons_Present = Find_Correct_Previous_Year_Button_return_xpath(driver)
    
    Output = pd.DataFrame()
    
    if No_Buttons_Present == True and Previous_Button_Present == False:
        # Like Anup Engineering
        Output = pd.concat([Output,Only_read_table_from_page(driver)],axis=1)
        Output_filename=company+' '+'Q_Report_'+dt.datetime.today().strftime('%Y_%m_%d_%H_%M_%S')+'.csv'
        Output_path = os.path.join(path_save,Output_filename)
        Output.to_csv(Output_path, sep=',',encoding='utf-8', index=False)
        print('Quarterly report data collected and saved for {}'.format(company))
    
    while True:
                             #/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[2]/td/center/font
                             #/html/body/div[4]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[2]/td/center/font
        No_more_data_xpath = '/html/body/div[2]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[2]/td/center/font'
        No_more_data_text = ' '
        try:
            No_more_data_text = driver.find_element_by_xpath(No_more_data_xpath).text.lower()
        except:
            pass
        if 'data' in No_more_data_text and 'not' in No_more_data_text and 'available' in No_more_data_text:
            # Save & break
            Output_filename=company+' '+'Q_Report_'+dt.datetime.today().strftime('%Y_%m_%d_%H_%M_%S')+'.csv'
            Output_path = os.path.join(path_save,Output_filename)
            Output.to_csv(Output_path, sep=',',encoding='utf-8', index=False)
            print('Quarterly report data collected and saved for {}'.format(company))
            break
        
        Next_Button_Present, Previous_Button_Present, Both_Next_Previous_Present, No_Buttons_Present = Find_Correct_Previous_Year_Button_return_xpath(driver)
        
        #Reading current page
        Output = pd.concat([Output,Only_read_table_from_page(driver)],axis=1)
        # Click on the correct 'previous' button
        Cancel_Adv_If_Present(driver)
        Correct_Previous_Button_Click(driver, Next_Button_Present, Previous_Button_Present, Both_Next_Previous_Present)
        # Input will be driver, Next_Button_Present, Previous_Button_Present, Both_Next_Previous_Present
        
        if Previous_Button_Present == False and Next_Button_Present == True:
            Output = pd.concat([Output,Only_read_table_from_page(driver)],axis=1)
            Output_filename=company+' '+'Q_Report_'+dt.datetime.today().strftime('%Y_%m_%d_%H_%M_%S')+'.csv'
            Output_path = os.path.join(path_save,Output_filename)
            Output.to_csv(Output_path, sep=',',encoding='utf-8', index=False)
            print('Quarterly report data collected and saved for {}'.format(company))
            break
        
###############################################################################
############################################################################### 
            
        
###############################################################################
###############################################################################
############## Function for getting quarterly balance sheet data ##############
def getQuaterlyReportDataForAllCompanies(driver, Company_Name_List, Ticker_Name_List, Company_found, Company_not_found, listname, path_save):
    for counter_1, company in enumerate(Company_Name_List):
        search_bar, search_bar_exist = Search_Bar_Check(driver)
        if search_bar_exist == True:
            Company_T_F = Typing_Checking_from_moneycontrol_search_bar(driver, counter_1, company, Ticker_Name_List, Company_found, Company_not_found)
            if Company_T_F == False:
                print('{} records not in moneycontrol.com, skipping to next company'.format(company))
                continue
            else:
                Quarterly_Report_Page = Navigate_to_Quarterly_reports(driver, company, counter_1, Ticker_Name_List)
                if Quarterly_Report_Page == True:
                    print("Reached {}'s quarterly reports page, collecting data...".format(company))
                    # Start Recording data, & save output
                    Record_Save_company_Quarterly_Record(driver, company, path_save)
                else:
                    print("Couldn't reach quarterly reports page of {}, skipping to next company".format(company))
                    continue
        else:
            print('Search bar of moneycontrol not found while searching for {}'.format(company))
            print('Consider re-evaluating the program & internet connection')
            break
        if counter_1 == len(Company_Name_List) - 1:
            print('Finished scraping over entire {}'.format(listname))
            break
        
###############################################################################
###############################################################################
            
        
###############################################################################
###############################################################################
################## Function for clean stopping of the driver ##################
def stopFirefoxDriver(driver):
    try:
        print('Shutting down Firefox Driver')    
        driver.quit()
        print('Driver Terminated successfully')
    except Exception as e:
        print('Error encountered while shutting down Firefox driver. \n {}'.format(e))
        
###############################################################################
###############################################################################
        
        
###############################################################################
###############################################################################
########################## Opening Moneycontrol.com ###########################
def To_Money_Control_Website(driver):
    moneycontrol_url = 'https://www.moneycontrol.com/'
    driver.get(moneycontrol_url)
    time.sleep(Time_Multiple*np.random.uniform(4.5, 5.5, size = 1))
    
    # Checking for search bar of moneycontrol.com
    search_bar, search_bar_exist = Search_Bar_Check(driver)
    return (search_bar, search_bar_exist)
    
###############################################################################
###############################################################################