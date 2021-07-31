from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials



#user input
eng=input("Engineer name: ")
nms=input("enter AD usernam: ")
pwd=input("enter AD password: ")



# use creds to create a client to interact with the Google Drive API
scope = ['https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive','https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\SwagathRamakrishnaBh\\Downloads\\client_secret.json", scope)
client = gspread.authorize(creds)
sheet = client.open("MyDataSheetPersonal")


#declaration
TicketDict = {}
highno=0


def color_row(work_sheet,next_row,r,g,b):
    next_row=str(next_row)
    
    work_sheet.format(f'A{next_row}:L{next_row}', {
             "backgroundColor": {
                 "red": r,
                 "green": g,
                 "blue": b
                },
               "horizontalAlignment": "CENTER",
               "textFormat": {
                  "foregroundColor": {
                    "red": 0,
                    "green": 0,
                    "blue": 0
                 },
                  "fontSize": 10,
                  "bold": False
                 }
              })
        

def next_available_row(worksheet):
     all = worksheet.get_all_values()
     next_row = len(all) + 1
     return next_row

def current_row(worksheet):
     all = worksheet.get_all_values()
     next_row = len(all) 
     return next_row

worksheets=sheet.worksheets()

branchlist={}

for  wsheet in worksheets:
    row=current_row(wsheet)
    ss=wsheet.acell(f'C{row}').value
    ss=str(int(ss)+1)
    branchlist[wsheet.title]=ss
    
    


#predifined

path="https://phx-p10y-jenkins-harbinger-prod-6.p10y.eng.nutanix.com/view/Master-LCC/"
driver = webdriver.Chrome('C:\\Users\\SwagathRamakrishnaBh\\Downloads\\chromedriver_win32\\chromedriver.exe')
wait=WebDriverWait(driver, 30)




#functions used
def CheckExistsByXpath(driver, xpath):
    try:
        path = driver.find_element_by_xpath(xpath)
    except NoSuchWindowException:
        return False
    except WebDriverException:
        return False
    except StaleElementReferenceException:
        return False
    return True

def FindHighBhuildNo():
    high_no= wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="buildHistory"]/div[2]/table/tbody/tr[2]/td/div[1]/a'))).get_attribute('innerText')
    high_no=high_no.lstrip('#')
    high_no=(high_no.encode('ascii', 'ignore')).decode("utf-8")
    return int(high_no) 

def get_run_date(run_date):
    r=run_date.split()
    print(r)
    print('\n')
    d=r[3].rstrip(',')+' '
    m=r[2].lstrip('(')+' '
    y=r[4]
    rundate=d+m+y
    return rundate

def login_jira():
    driver.get('https://jira.nutanix.com/')
    username=wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="okta-signin-username"]')))
    username.send_keys(nms)
    password=wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="okta-signin-password"]')))
    password.send_keys(pwd)
    signin=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="okta-signin-submit"]')))
    signin.send_keys(Keys.ENTER)
    sendpush=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="form8"]/div[2]/input')))
    sendpush.send_keys(Keys.ENTER)
    driver.switch_to.window(driver.window_handles[0])

def getsingleticket(impropertickets):
   
    properticket=re.search(r'[D]{1}[I]{1}[A]{1}[L]{1}[-][0-9]{4}', impropertickets,re.I)
    return(properticket.group().upper())
    

def getticketdetails(ticket):
    failuretype=''
    faileddueto=''
    driver.switch_to.window(driver.window_handles[1])
    searchbar=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="quickSearchInput"]')))
    searchbar.send_keys(ticket)
    searchbar.send_keys(Keys.ENTER)
    summary=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="summary-val"]'))).get_attribute("innerText")
    type=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="type-val"]'))).get_attribute("innerText")
    type=type.strip()
   
    if type=='Functional Test (Testbody)':
        failuretype='Test Failure'
        faileddueto='Product'

    elif type=='Compile':
        failuretype='Build Failure'
        faileddueto='Product'
    
    elif type=='Tool Issue':
        failuretype='Build Failure'
        faileddueto='Tools-Infra'

    elif type=='Product Deployment':
        failuretype ='Deployment Failure'
        faileddueto='Product'

    elif type=='Product Issue':
        failuretype ='Deployment Failure'
        faileddueto='Product'

    elif type=='Unit Test(Ptest)':
        failuretype ='ptest Failure'
        faileddueto='Product'

    elif type=='IT Issue':
        failuretype ='Deployment Failure'  
        faileddueto='IT-Issue'
    
    ticketlist=list()
    ticketlist.append(summary)
    ticketlist.append(failuretype)
    ticketlist.append(faileddueto)
    TicketDict[ticket]=ticketlist
    return ticketlist


driver.get(path)
driver.maximize_window()

#open new window for jira
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[1])
login_jira()

for i in branchlist:

        #complete branch details
        branchdetail=list()



        branch=wait.until(EC.presence_of_element_located((By.LINK_TEXT,"LCC » "+i)))
        action=ActionChains(driver)
        action.click(on_element=branch)
        action.perform()

        if CheckExistsByXpath(driver,'//*[@id="buildHistory"]/div[2]/table/tbody/tr[2]/td/div[1]/a'):
            highno=FindHighBhuildNo()
            
        else:
            highno=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="buildHistory"]/div[2]/table/tbody/tr[3]/td/div[1]/a'))).get_attribute('innerText')
            highno=highno.lstrip('#')
            highno=(highno.encode('ascii', 'ignore')).decode("utf-8")
            highno=int(highno)

        #run=wait.until(EC.presence_of_element_located((By.LINK_TEXT,'#'+branchlist[i])))
        #action=ActionChains(driver)
        #action.click(on_element=run)
        #action.perform()


        progress=False


        checkhighno=int(branchlist[i])
        
        while progress==False or checkhighno<=highno:
            
            if checkhighno>highno or progress==True:
                break
           
            runlist=list()
            flag='Green Run (Pass)'
            
            #click on build no
            run=wait.until(EC.presence_of_element_located((By.LINK_TEXT,'#'+branchlist[i])))
            action=ActionChains(driver)
            action.click(on_element=run)
            action.perform()

            

            progress=CheckExistsByXpath(driver,'//*[@id="main-panel"]/h1/div/table/tbody/tr/td[1]')
            
            if progress==True:
                break
            
            
            
            #get date
            if CheckExistsByXpath(driver,'//*[@id="tasks"]/div[15]/span/a/span[2]'):
                run_date=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="main-panel"]/h1'))).get_attribute('innerText')
                rundate= get_run_date(run_date)
                runlist.append(rundate)
            else:
                run=wait.until(EC.presence_of_element_located((By.LINK_TEXT,'#'+branchlist[i])))
                action=ActionChains(driver)
                action.click(on_element=run)
                action.perform()
                run_date=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="main-panel"]/h1'))).get_attribute('innerText')
                rundate= get_run_date(run_date)
                runlist.append(rundate)
             
            runlist.append(eng)
            runlist.append(int(branchlist[i]))
            

            # get duration
            duration=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="main-panel"]/div[1]/div[2]/a'))) 
            durations=duration.text                                         
            
            #get status of run
            status=wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[4]/div[2]/h1/span/*[name()='svg']/*[name()='use']")))
            statuses=status.get_attribute("href")
            statuses=statuses.split('#')
           
            if statuses[1]=='last-failed':
                ticketno=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="description"]/div[1]'))) 
                
                maybemultipleticket=ticketno.text
                if len(maybemultipleticket)!=0:
                    ticketnos=getsingleticket(maybemultipleticket)  
                
                    if ticketnos in TicketDict:
                        flag=TicketDict[ticketnos][1] #failure type
                        runlist.append(flag)
                        runlist.append('')
                        runlist.append('')
                        runlist.append(TicketDict[ticketnos][2]) #failed due to
                        runlist.append('') 
                        runlist.append(ticketnos)
                        runlist.append(durations)
                        runlist.append(TicketDict[ticketnos][0]) #summary
                        runlist.append('')
                    else:
                        ticketdetail=getticketdetails(ticketnos)
                        flag=ticketdetail[1]
                        runlist.append(flag)
                        runlist.append('')
                        runlist.append('')
                        runlist.append(ticketdetail[2])
                        runlist.append('')
                        runlist.append(ticketnos)
                        runlist.append(durations) 
                        runlist.append(ticketdetail[0])
                        runlist.append('')

                    driver.switch_to.window(driver.window_handles[0])    
                else:
                    break   
                    
               

            elif statuses[1]=='last-successful':
               
                flag='Green Run (Pass)'
                runlist.append(flag)
                runlist.append('')
                runlist.append('')
                runlist.append('')
                runlist.append('')
                runlist.append('')
                runlist.append(durations)
                #runlist.append('')
                #runlist.append('')
            else:
                desc=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="description"]/div[1]'))) 
                desc=desc.text
                flag='Aborted' 
                runlist.append(flag)
                runlist.append('')
                runlist.append('')
                runlist.append('')
                runlist.append('')
                runlist.append('')
                runlist.append(durations)
                runlist.append('')
                if len(desc)!=0:
                    runlist.append(desc)
                else:
                    runlist.append('')
           


            #update run data on branch list
            branchdetail.append(runlist)

            branchlist[i]=int(branchlist[i])
            branchlist[i]=branchlist[i]+1
            checkhighno=branchlist[i]
            branchlist[i]=str(branchlist[i])
            
            #go back by clicking branch-lcc
            branchlcc=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="breadcrumbs"]/li[7]/a')))
            action=ActionChains(driver)
            action.click(on_element=branchlcc)
            action.perform()
            


       

        sheets = client.open("MyDataSheetPersonal")
        worksheet=sheets.worksheet(i)
        next_row = next_available_row(worksheet)
        for j in range(len(branchdetail)):
            
            worksheet.insert_row(branchdetail[j], next_row, value_input_option='USER_ENTERED') 
            worksheet.update(f'A{next_row}',branchdetail[j][0],value_input_option='USER_ENTERED')
            if branchdetail[j][3]=='Green Run (Pass)':
                color_row(worksheet,next_row,0.576, 0.769,0.49)
            elif branchdetail[j][3]=='Test Failure' or  branchdetail[j][3]=='ptest Failure':
                color_row(worksheet,next_row,1.00,0.851,0.40)
            elif branchdetail[j][3]=='Deployment Failure':
                color_row(worksheet,next_row,0.94,0.28,0.28)
            elif branchdetail[j][3]=='Aborted':
                color_row(worksheet,next_row,0.93,0.93,0.93)
            else:
                color_row(worksheet,next_row,0.643,0.761,0.957)
            next_row=next_row+1
            driver.implicitly_wait(5)
        #go to master-lcc
        masterlcc=wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="breadcrumbs"]/li[3]/a')))
        action=ActionChains(driver)
        action.click(on_element=masterlcc)
        action.perform()

#teardown
driver.close()
driver.quit()

