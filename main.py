#################################################[ IMPORTS/SETTINGS/GLOBAL VARIABLES ] #################################################
#import all constants required
from constants import splash, colorize, constant_menu, colors

#import os to manage files / find paths / create files and folders
import os

#import CSV to process CSV files (PRIMS & ADHOC)
import csv

#import time to manage timings using time.sleep() - THIS NEEDS TO BE ADRESSED USING SELENIUM SPECIFIC EXCEPTION HANDLING WAITS
import time

#import datetime to compare date of last eval from BOL with ETJ Qualification/Certification Data. Import timedelta to add 1 day after the last eval to create new block 14.
import datetime
from datetime import timedelta

# importing shutil module  
import shutil  

#import webridver & json to set appstate
from selenium import webdriver
import json

#import pyodbc to manage Access database
import pyodbc

#import ConfigParser to process config located in /config/config.ini
from configparser import ConfigParser

#define global variables
#those are the ones that are groupped up in either FLTMPS or PRIMS/OMPF and variables are used to hold the entire string which later is sprlit accordingly
name = ''
name_folder = ''
parent_activity = ''
pending_activity = ''
projected_rotation = ''
comments_pitch = '10 POINT'

#those are variables that are written to the database. Comments next to them reflect column name so I don't have to re-open it in Access every time
block1_name = ''                                    #Name
block2_rate = ''                                    #Rate
block3_designation = ''                             #Desig
block4_ssn = ''
block5a_act = ''
block5b_fts = ''
block5c_inact = ''
block5d_at_adsw_265 = ''
block6_uic = ''
block7_station = ''
block8_status = ''
block9_date_reported = ''
block10_periodic = ''
block11_doa = ''
block12_promotion = ''
block13_special = ''
block14_from = '0'
block15_to = ''
block16_nob = ''
block17_regular = ''
block18_concurent = ''
block20_pfa = ''
block21_bilet_subcategory = ''
block22_rep_senior_name = ''
block23_rep_senior_grade = ''
block24_rep_senior_desig = ''
block25_rep_senior_title = ''
block26_rep_senior_uic = ''
block27_rep_senior_ssn = '000000000'
block28_command_empl = ''
block29a_primary_duty = ''
block29a_duty = ''
block30_date_consueled = ''
block31_consuelor = ''
block33_professional_knowledge = ''
block34_quality_of_work = ''
block35_command_involvement = ''
block36_mil_bearing = ''
block37_initiative = ''
block38_teamwork = ''
block39_leadership = ''
block42_rater = ''
block43_writeup = ''
block44_quals = ''
block45_individual = ''
block46_summary = ''
block47_retention = ''
block49_senior_rater = ''

#set app state to chrome driver so it prints to PDF automatically when asked to print
appState = {
"recentDestinations": [
    {
        "id": "Save as PDF",
        "origin": "local",
        "account": ""
    }
],
"selectedDestinationId": "Save as PDF",
"version": 2
}

profile = {'printing.print_preview_sticky_settings.appState': json.dumps(appState)
}

#create all of the options
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', profile)
options.add_argument('--kiosk-printing')
options.add_argument("--print-to-pdf")
#those two will prevent cert error (so the CA certs don't have to be installed)
options.add_argument('--ignore-ssl-errors=yes')
options.add_argument('--ignore-certificate-errors')
#this is here to clear up console output, only "fatal" logs will be displayed
options.add_argument("--log-level=3")

#**************************************** MAIN FUNCTIONS ****************************************# 
#! THIS IS WHERE I INITIATE DRIVER AND INIT CONFIGURATION

#driver init
def driver_init():
    #driver defined as global so I don't have to spawn it every time if it's already spawned
    global driver
    #spawn instance of webdriver with options defined above
    driver = webdriver.Chrome('lib/chromedriver.exe', options=options)
    driver.set_window_size(800, 600)

#init config for ETJ
def driver_init_etj_blank():
    driver.get('https://ntmpsweb.ncdc.navy.mil/etjclient/')
    time.sleep(3) #this shit is super random and needs to be adjusted using selenium exception waits
    driver.switch_to.alert.accept()
    driver.find_element_by_id('btnCACLogin').click()
    
#init config for BOL
def driver_init_bol_blank():
    driver.get('https://www.bol.navy.mil/bam/')
    time.sleep(3) #this shit is super random and needs to be adjusted using selenium exception waits
    driver.switch_to.alert.dismiss()
    driver.find_element_by_id('CACLoginButton').click()

#function to create directories needed
def dir_maker(name_folder, switch):
        #get path to desktop and create folder with the name of evaluation if folder does not already exist
    if not os.path.exists(name_folder):
        os.makedirs(name_folder)
        os.makedirs(name_folder+switch)
#if subfolder BOL/ETJ does not exist withing folder name on desktop, create one        
    elif not os.path.exists(name_folder+switch):
        os.makedirs(name_folder+switch)

#! THIS IS WHERE I FETCH EVALS

def fetch_evals():
    #get name figured out and create paths required
    name = driver.find_element_by_id('PageHeader_ucbolHeader_UserBanner1_lblName').text
    name_folder = os.path.expanduser('~\\Desktop\\{} {} {}'.format((name.split())[0], (name.split())[1], (name.split())[2]))

    #create directories needed
    dir_maker(name_folder, '\\BOL')

    #navigate all the way to OMPF
    driver.find_element_by_id('MenuItemRepeater_ctl17_MenuItemLink').click()
    driver.find_element_by_id('YesButton').click()
    global block4_ssn
    block4_ssn = '00000' + driver.find_element_by_id('UserDescriptionLabel').text.split('/')[0]
    driver.find_element_by_xpath('//*[@id="TabStrip"]/div/ul/li[2]/a').click()

    #set and empty list named document_adress where we store resource key used by OMPF to access the right PDF
    document_address = []

    #select the table we want to iterate through (named ResultsGrid_ctl00)
    table_id = driver.find_elements_by_id('ResultsGrid_ctl00')[0]

    global block14_from
    #find all rows and skip the first one since it's a header and doesn't have columns
    rows = table_id.find_elements_by_tag_name('tr')[1:]
    for row in rows:
        #set empty list named row_value where we add text value of each cell as another list item which we access later
        row_value = []
        #find all columns
        columns = row.find_elements_by_tag_name('td')      
        for column in columns:
            row_value.append(column.text)
        #if we find a row which has a value of 1616/26 (eval) in the 3rd column 
        if row_value[2] == '1616/26':
            #get the value of the value tag of the first invisible (attached to the checkbox) column
            document_address.append(driver.find_element_by_xpath('//*[@id="'+row.get_attribute('id')+'"]/td[1]/input').get_attribute('value'))
            #while we're here let's check the date of that eval and if it's greater (newer) than any of the others set it as the last eval date
            if int(row_value[6]) > int(block14_from):
                block14_from = row_value[6]
    #correct FROM date from YYYYMMDD format to MM/DD/YYYY format that's accepted by Access (this is stupid and I don't like it)
    #assign value of the last eval to the block 14 and convert it to datetime object using format YYYYMMDD
    block14_from = datetime.datetime.strptime(block14_from, '%Y%m%d')
    #add an extra day since the block14 FROM is the day after the last regular eval.
    block14_from = block14_from + datetime.timedelta(days = 1)
    #extract just the date from it and convert it to MM/DD/YYYY required by access
    block14_from = block14_from.date().strftime('%m/%d/%Y')

    #now that we have all evals keys (document addresses of all evals) we check how many evals we have stored
    #if we have more than 3 evals set number_of_evals to 3 only (since we only need last 3 evals), if it's smaller set number_of_evals to the length of the list 
    if len(document_address) >= 3:
        number_of_evals = 3
    else:
        number_of_evals = len(document_address)

    #iterate through all eval keys, access documents and print them (sleeper is needed since window.print() is slow)
    for address in range(0,number_of_evals):
            driver.get('https://www.bol.navy.mil/OMPF/Document.aspx?resource='+str(document_address[address]))
            driver.execute_script("window.print();")
            time.sleep(10)

    #move all to desktop folder & rename accordingly
    shutil.move((os.path.expanduser('~\\Downloads\\Document.aspx.pdf')), name_folder+'\\BOL\\EVAL 1.pdf')
    shutil.move((os.path.expanduser('~\\Downloads\\Document.aspx (1).pdf')), name_folder+'\\BOL\\EVAL 2.pdf')
    shutil.move((os.path.expanduser('~\\Downloads\\Document.aspx (2).pdf')), name_folder+'\\BOL\\EVAL 3.pdf')

#! THIS IS WHERE I FETCH PRIMS

def fetch_prims():
    block14_from = '3/16/2019'
    
    #get name figured out and create paths required
    name = driver.find_element_by_id('PageHeader_ucbolHeader_UserBanner1_lblName').text
    name_folder = os.path.expanduser('~\\Desktop\\{} {} {}'.format((name.split())[0], (name.split())[1], (name.split())[2]))
    
    #create directories needed
    dir_maker(name_folder, '\\BOL')

    #navigate to the PFA report for all cycles
    driver.find_element_by_id('MenuItemRepeater_ctl19_MenuItemLink').click()
    driver.find_element_by_partial_link_text('Member').click()
    time.sleep(5)
    driver.find_element_by_partial_link_text('Reports').click()     
    driver.find_element_by_partial_link_text('PFA Listing All Cycles').click()    
    time.sleep(10)
    #since report is opened in new tab switch focus to that tab
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_xpath('//*[@id="ajaxReportViewer1"]/div/div[2]/div[2]/ul[1]/li[10]/a').click()
    time.sleep(1)
    #download report in PDF format    
    driver.find_element_by_partial_link_text('Acrobat (PDF) file').click()    
    time.sleep(5)
    #download report in CSV format so we can extract last PFA cycles
    driver.find_element_by_xpath('//*[@id="ajaxReportViewer1"]/div/div[2]/div[2]/ul[1]/li[10]/a').click()
    time.sleep(1)
    driver.find_element_by_partial_link_text('CSV (comma delimited)').click()    
    time.sleep(5)
    #close report tab
    driver.close()
    #switch focus back to the main tab
    driver.switch_to.window(driver.window_handles[0])
    #move and rename file accordingly 
    shutil.move((os.path.expanduser('~\\Downloads\\MemberPFAListingAllCycles.pdf')), name_folder+'\\BOL\\PRIMS.pdf')
    shutil.move((os.path.expanduser('~\\Downloads\\MemberPFAListingAllCycles.csv')), name_folder+'\\BOL\\PRIMS_TEMP.csv')    

    #process PRIMS csv report to the data usable in block 20
    #lets create some dictionary to store cycle and result 
    prt_results = {}
    #open cvs file and spawn instance of csv reader

    with open(name_folder+'\\BOL\\PRIMS_TEMP.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        #for each row except header (header is identified as having 'textBox6' in the first column
        for row in csv_reader:
            if row[0] != 'textBox6':
                cycle = ''
                result = ''
                #if date of the PRT was later than the "to" date of the last eval - that means that those PRTs did not get included in last eval
                if datetime.datetime.strptime(row[5], '%m/%d/%Y') >= datetime.datetime.strptime(block14_from, '%m/%d/%Y'):
                    if row[7] == 'Participant':
                            cycle = '{}-{}'.format(row[1].split()[0][2:4], row[1].split()[1])
                            result = row[11]
                            print(cycle, result)
                            #here i need to split pass/faill/waived etc but I need more data\
                            prt_results.update( {cycle : result} )
                    elif row[7] == 'Excused':
                        cycle = '{}-{}'.format(row[1].split()[0][2:4], row[1].split()[1])
                        result = 'Excused'
                        print(cycle, result)
                        prt_results.update( {cycle : result} )
    print('Block 20:')
    print(prt_results)
    input('press enter')  

#! THIS IS WHERE I FETCH ETJ DATA

#!fetching admin page
def fetch_admin():
    #get name figured out and create paths required
    global name
    name = driver.find_element_by_id('ContentPlaceHolder1_lblName').text
    global name_folder
    name_folder = os.path.expanduser('~\\Desktop\\{} {} {}'.format((name.split())[1], (name.split())[2], (name.split())[3]))
    global block1_name
    block1_name = (name.split())[1] + ', ' + (name.split())[2] + ' ' + ((name.split())[3])[0]
    global block2_rate
    block2_rate = (name.split())[0]
    ocupation = driver.find_element_by_id('ContentPlaceHolder1_lblOccupationCatDesc').text
    global block3_designation
    block3_designation = ((ocupation.split('('))[1]).split(')')[0]
    parent_activity = driver.find_element_by_id('ContentPlaceHolder1_lblParentUIC').text
    #this is here to populate portion of block 43 and adjust block 15 accordingly if member transfers before eval cycle.
    pending_activity = driver.find_element_by_id('ContentPlaceHolder1_lblPendingUIC').text
    projected_rotation = driver.find_element_by_id('ContentPlaceHolder1_lblPRD').text
    global block6_uic
    block6_uic = parent_activity.split()[0][1:]
    global block7_station
    block7_station = parent_activity.split()[1] + ' ' + parent_activity.split()[2]
    global block9_date_reported
    block9_date_reported = driver.find_element_by_id('ContentPlaceHolder1_lblDateRcvd').text

    #create directories needed
    dir_maker(name_folder, '\\ETJ')

    #navigate to ETJ Administrative Data and print, rename to ETJ Admin and move to the folder created on desktop
    driver.execute_script("window.print();")
    shutil.move((os.path.expanduser('~\\Downloads\\ETJ.pdf')), name_folder+'\\ETJ\\ETJ Administrative Data.pdf')

#!fetching awards page
def fetch_awards():
    #get name figured out and create paths required
    name = driver.find_element_by_id('ContentPlaceHolder1_lblName').text
    name_folder = os.path.expanduser('~\\Desktop\\{} {} {}'.format((name.split())[1], (name.split())[2], (name.split())[3]))

    #get path to desktop and create folder with the name of evaluation if folder does not extist if it does delete content
    dir_maker(name_folder, '\\ETJ')

    #navigate to ETJ Awards and print, rename to ETJ Awards and move to the folder created on desktop
    driver.find_element_by_xpath('//*[@id="ListTopMenu"]/li[8]/a').click()
    driver.execute_script("window.print();")
    shutil.move((os.path.expanduser('~\\Downloads\\ETJ.pdf')), name_folder+'\\ETJ\\ETJ Awards Datas.pdf')

#!fetching quals/certs page
def fetch_quals():
    #get name figured out and create paths required
    name = driver.find_element_by_id('ContentPlaceHolder1_lblName').text
    name_folder = os.path.expanduser('~\\Desktop\\{} {} {}'.format((name.split())[1], (name.split())[2], (name.split())[3]))

    #create directories needed
    dir_maker(name_folder, '\\ETJ')

    #navigate to ETJ Quals/Certs and print, rename to ETJ Quals and move to the folder created on desktop
    driver.find_element_by_xpath('//*[@id="ListTopMenu"]/li[7]/a').click()
    driver.execute_script("window.print();")
    shutil.move((os.path.expanduser('~\\Downloads\\ETJ.pdf')), name_folder+'\\ETJ\\ETJ Quals & Certs Data.pdf')

#! THIS IS WHERE WE PROCESS PRIMS CSV REPORT


#! THIS IS WHERE WE WRITE DATABASE
def database_writer():
    #get current working directory and get con
    cwd = os.getcwd()
    #spawn instance of ConfigPraser and read config placed in /config/config.ini
    config = ConfigParser()
    config.read(cwd + '\\config\\config.ini')

    input('PRESS ENTER TO WRITE DATABASE')
    print(block1_name)
    print(block2_rate)
    print(block3_designation)
    print(block4_ssn)
    print(block6_uic)
    print(block7_station)

    if block2_rate[-1] == '1':
        block_senior_rater = config.get('e6','senior_rater')
        block_rater = config.get('e6','rater')
        block_reporting_senior = config.get('e6','reporting_senior')
        block15_to = config.get('e6','to')
    elif block2_rate[-1] == '2':
        block_senior_rater = config.get('e5','senior_rater')
        block_rater = config.get('e5','rater')
        block_reporting_senior = config.get('e5','reporting_senior')
        block15_to = config.get('e5','to')
    elif block2_rate[-1] == '3':
        block_senior_rater = config.get('e4','senior_rater')
        block_rater = config.get('e4','rater')
        block_reporting_senior = config.get('e4','reporting_senior')
        block15_to = config.get('e4','to')
    else:
        block_senior_rater = config.get('airman','senior_rater')
        block_rater = config.get('airman','rater')
        block_reporting_senior = config.get('airman','reporting_senior')
        block15_to = config.get('e3','to')

    print(block_senior_rater)
    print(block_rater)
    print(block_reporting_senior)
    input()

    #I don't like this portion but here is a simple explanation why it had to be done that way. NAVFIT98 is all kinds of messed up. It uses default MDB database (prior Access 2007). with Access 2007, Microsoft introduced a new database engine (the Office Access Connectivity Engine, referred to also as ACE and Microsoft Access Database Engine) to replace the Jet database used for MDB files. The problem is that NAVFIT98 writers did not re-write the code to accomodate the changes. Instead they're hiding mdb database under accbd extension. For that reason blank .mdb database is moved over and worked on, then extension name is changed to accdb. It makes no difference for neither MS Access nor NAVFIT98 but it does for Microsoft Access Database Engine.
    
    #get current working directory
    cwd = os.getcwd()
    shutil.copy((cwd + '\\lib\\blank.mdb'), name_folder)
    db_location = name_folder+'\\blank.mdb'

    #initiate ODB driver with right parameters
    database_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'

    conn = pyodbc.connect(driver = database_driver, dbq = db_location, autocommit = True)
    cursor = conn.cursor()
    cursor.execute('select * from Reports')

    sql = 'UPDATE [Reports] SET [Name] = ?, [Rate] = ?, [Desig] = ?, [SSN] = ?, [Active] = ?, [UIC] = ?, [ShipStation] = ?, [PromotionStatus] = ?, [DateReported] = ?, [Periodic] = ?, [FromDate] = ?, [ToDate] = ? WHERE [ReportType] = ?'
    params = (block1_name, block2_rate, block3_designation, block4_ssn, '1', block6_uic, block7_station, 'REGULAR', block9_date_reported, '1', block14_from, block15_to, 'Eval')
    cursor.execute(sql, params)

    conn.commit()
    cursor.close()
    conn.close()
    #this one has to stay for the reason stated above. M$ Access Driver would crash since it's not real MADE but Jet database.
    os.rename(name_folder + '\\blank.mdb', name_folder + '\\'+ block2_rate + ' ' + (name.split())[1] + '.accdb')  
    input('DATABASE WRITTEN SUCCESSFULLY')

#**************************************** MAIN MENU ****************************************# 

#define menu items
menu_elements = ['(1) FetchFetch last 3 evals', '(2) Fetch last PRIMS ', '(3) Fetch BOL data (EVALS/PRIMS)', '(4) Fetch admin page', '(5) Fetch awards page', '(6) Fetch qualification/certification page', '(7) Fetch ETJ data (ADMIN/AWARDS/QUALS)', '(8) Fetch all Data (EVALS/PRIMS/ADMIN/AWARDS/QUALS)', '(9) Fetch all data (EVALS/PRIMS/ADMIN/AWARDS/QUALS) and create database']

#menu loop
def show_menu():     
    loop = True
    while loop:
        splash()
        for item in menu_elements:
            print(colorize(item,'pink'))
        constant_menu()
        choice = input ('Please make a choice: ')
        if choice == '1':
            print('Attempting to fetch EVALS ...')
            driver_init()
            driver_init_bol_blank()
            fetch_evals()
            driver.quit()
        elif choice == '2':
            print('Attempting to fetch PRIMS ...')
            driver_init()
            driver_init_bol_blank()
            fetch_prims()
            driver.quit()
        elif choice == '3':
            print('Attempting to fetch BOL data (PRIMS/EVALS)')
            driver_init()
            driver_init_bol_blank()
            fetch_evals()
            driver_init_bol_blank()
            fetch_prims()
            driver.quit()
        elif choice == '4':
            print('Attempting to fetch ETJ Admin page ...')
            driver_init()
            driver_init_etj_blank()
            fetch_admin()
            driver.quit()
        elif choice == '5':
            print('Attempting to fetch ETJ Awards page ...')
            driver_init()
            driver_init_etj_blank()
            fetch_awards()
            driver.quit()
        elif choice == '6':
            print('Attempting to fetch ETJ Quals/Certs page ...')
            driver_init()
            driver_init_etj_blank()
            fetch_quals()
            driver.quit()
        elif choice == '7':
            print('Attempting to fetch ETJ data ...')
            driver_init()
            driver_init_etj_blank()
            fetch_admin()
            fetch_awards()
            fetch_quals()
            driver.quit()
        elif choice == '8':
            print('Attempting to fetch ETJ and BOL data ...')
            driver_init()
            driver_init_bol_blank()
            fetch_evals()
            driver_init_bol_blank()
            fetch_prims()
            driver.quit()
            driver_init()
            driver_init_etj_blank()
            fetch_admin()
            fetch_awards()
            fetch_quals()
            driver.quit()
            database_writer()
        elif choice == '9':
            print('Attempting to fetch ETJ and BOL data and write it to database ...')
            driver_init()
            driver_init_bol_blank()
            fetch_evals()
            driver_init_bol_blank()
            fetch_prims()
            driver.quit()
            driver_init()
            driver_init_etj_blank()
            fetch_admin()
            fetch_awards()
            fetch_quals()
            driver.quit()
            database_writer()             
        elif choice == 'q':
            break            
        else:
            print('Not a valid option. Try again.')
            time.sleep(2)
    #say bye bye :)
    splash()    
    print(colorize('''
    
    Bye bye:) ''','blue'))
    time.sleep(2)