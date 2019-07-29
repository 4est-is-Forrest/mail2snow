#General
"""
Requires Chromedriver.exe to be in local directory and Chrome to be installed on Windows.
"""
#Problem
"""
My team's shared helpdesk inbox is a high volume. It was difficult to make email SLA (48 hours) even with 4-6 people ignoring helpdesk calls in order to process emails.

Much of the time working the box is occupied with merely filling in multiple fields on Catalog Tasks and Inicidents.

We needed the capability to process tickets into the ticketing system from email while keeping a human like accuracy in half a minute rather than five.

Limitations: ServiceNow API access and Web Service restrictions placed on me forced a solution involving Selenium and simulating an agent entering tickets in the UI.
"""

#Solution
"""
Extract assignment group, client, and short description from email's subject line.
Access ServiceNow with supported login and Chrome cookie managements.
Use a combination of encoded queries and Javascript to submit the email as a ticket.
ReplyAll to the end user with a generic message and ticket number.
"""

import win32com.client
import os,re,sys,time,random,json
from datetime import datetime  
from datetime import timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import pyperclip
import selenium.webdriver.support.ui as ui
from pathlib import Path
import shutil
import shelve
from fuzzywuzzy import fuzz

###############################################################################
#Constants
###############################################################################

#Directory to temporarily store emails for upload to ServiceNow
if not os.path.exists('Emails'):
    os.makedirs('Emails')

#Chrome Driver Options
CHROMEDRIVER_PATH= os.getcwd()+ '\\chromedriver.exe'
size= '1200,900'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--window-size={}'.format(size))
chrome_options.binary_location = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
chrome_options.add_argument("user-data-dir=" + os.getcwd() + "\\Driver Data")#Allows for cookie login

#Get Outlook API
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
folders= namespace.Folders
root_folder = folders['My.Helpdesk@foo.net']

exclude = ['Helpdesk boxes that should be excluded']

#Spell-Checking Resource - Download from my machine, if not stored locally on script user's machine.
emaildb = Path(r"Data\emaildb")
data_path = Path(r'\\my_machine\path\Data')
if not os.path.exists('Data'):
    shutil.copytree(data_path,'Data')
"""Unpack Database"""
with shelve.open(emaildb.as_posix()) as db:
    lqlist = db['lowercase_queues_list']

###############################################################################
#Javascript
###############################################################################

#Person running script SNOW sys_id
def current_snow_user():
    g_user = None
    while g_user == None:
        try:
            x = driver.execute_script('{return g_user.userID};')
            g_user = x
        except:
            time.sleep(.3)
    return g_user

#Client sys_id
def core_company(x):
    result = None
    while result == None:
        try:
            driver.execute_script("window.rec = new GlideRecord('core_company');")
            time.sleep(.5)
            driver.execute_script("window.rec.addQuery('name',\'" + x + "\');function callback(rec) {while (rec.next()) {y = rec.sys_id}};window.rec.query(callback);")
            time.sleep(.5)
            result = driver.execute_script("return window.rec.sys_id")
        except:
            time.sleep(.5)
    return result

#Generic Client User
def user(x):
    if x == '<Client>':
        user = 'Generic <Client>'
    else:
        user = 'Generic User {}'.format(x)
    result = None
    while result == None:
        try:
            driver.execute_script("window.rec = new GlideRecord('sys_user');")
            time.sleep(.5)
            driver.execute_script("window.rec.addQuery('name',\'" + user + "\');function callback(rec) {while (rec.next()) {y = rec.sys_id}};window.rec.query(callback);")
            time.sleep(.5)
            result = driver.execute_script("return window.rec.sys_id")
        except:
            time.sleep(.5)
    return result

#Generic Client Config Item
def cmdb_ci(company):
    result = None
    while result == None:
        try:
            driver.execute_script("window.rec = new GlideRecord('cmdb_ci');")
            time.sleep(.5)
            driver.execute_script("window.rec.addQuery('name',\'Generic_" + company + "\');function callback(rec) {while (rec.next()) {y = rec.sys_id}};window.rec.query(callback);")
            time.sleep(.5)
            result = driver.execute_script("return window.rec.sys_id")
        except:
            time.sleep(.5)
    return result

#Assignment Group sys_id
def group(x):
    result = None
    while result == None:
        try:
            driver.execute_script("window.rec = new GlideRecord('sys_user_group');")
            time.sleep(.5)
            driver.execute_script("window.rec.addQuery('name',\'" + x + "\');function callback(rec) {while (rec.next()) {y = rec.sys_id}};window.rec.query(callback);")
            time.sleep(.5)
            result = driver.execute_script("return window.rec.sys_id")
        except:
            time.sleep(.5)
    return result

###############################################################################
#Functions
###############################################################################

#Reset mail: QOL function
def reset():
    msg.subject = re.sub('Working\.\.\.','',msg.subject)
    msg.unread = True
    msg.save()

#Short-hand WebDriverWait: QOL function
def wait(x,y,z):
    els = [By.XPATH, By.CSS_SELECTOR, By.PARTIAL_LINK_TEXT]
    return WebDriverWait(driver, x).until(EC.presence_of_element_located((els[y], z)))

#Deal with persistent logout page redirection after session expires.
def garbage():
    wait(10,0,'/html/body')
    if 'logout' in driver.current_url:
        sys.exit()

#Login Handler
def login():
    driver.get('https://<instance>.service-now.com/nav_to.do')
    wait(10,0,'/html/body')
    time.sleep(1)
    garbage()
    if 'nav_to.do?' in driver.current_url:
        wait(15,0,'//*[@id="filter"]')
        pass
    elif '?id=sso&portal-id=null' in driver.current_url:
        driver.get('https://<instance>.service-now.com/login_with_sso.do?glide_sso_id=<id>')
        wait(7,0,'/html/body')
        login()
        garbage()
    elif 'desired login page' in driver.current_url:
        url = driver.current_url
        while driver.current_url == url:
            time.sleep(.5)
        garbage()
        wait(45,0,'//*[@id="filter"]')

#Queue Spellchecker
def queue_spellcheck(q):
    try:
        temp = q.lower()
        if q not in lqlist:
            cur = 0
            for i in lqlist:
                if fuzz.partial_ratio(temp,i) > 80 and fuzz.partial_ratio(temp,i) > cur:
                    q = i
                    cur = fuzz.partial_ratio(temp,i)
        return q.lower()
    except:
        return q.lower()

#Submit email as ServiceNow INC ticket
def ticket(m):

    #Get source email address
    try:
        email = re.findall(r'&&\s*(.+)\s*&&',m.subject)[0]#Optional argument that user can place in subject line to override who is believed to be the sender.
    except:                    
        email= m.senderemailaddress
        if '@' not in email:
            email= m.sender.GetExchangeUser().PrimarySmtpAddress

    #Ticket body to be pasted
    ticketBody= 'From: ' + email + '\nSent: ' + datetime.strftime(m.receivedtime, '%Y-%m-%d %H:%M:%S') + '\nTo: ' + m.to + '\nCC: ' + m.cc + '\nSubject: ' + m.subject + '\n\n' +  m.body
    
    m.Unread = False #Don't interferre with concurrent script instances.

    #Check for assignment, spellcheck
    try:
        queue = re.findall('\{\{(\S+)\}\}',m.subject,re.I)[0]
    except:
        m.subject= 'No Assignment Group - ' + m.subject
        m.save()
    queue = queue_spellcheck(queue)

    #Define short description, defined optionally
    try:
        short = re.findall('\$\$\s*(.+)\s*\$\$',m.subject,re.I)[0]
    except:
        short = re.sub('[\$\$\{\{%%&&].+','',m.subject)

    #Define client, "[generic]" if None
    try:
        company = re.findall('%%\s*(.+)\s*%%', m.subject)[0]
    except:
        company = '[generic]'

    #Print arguements found for user
    print('\n{}\n\n{}\n\n{}\n\n{}'.format(m.subject,email,m.receivedtime,queue.upper()))
    print('=' * 50)

    #Record in order to update shelve object at a later date.
    with open('data.csv','a+') as doc:
        doc.write('{},{},{}\n'.format(email,company,queue))

    #Mark in progress for other humans
    m.subject = 'Working...'+ m.subject
    m.save()

    #Params for URL - JS functions
    comp_id = core_company(company)
    user_id = user(company) #Tickets this account are opened under client or projects, not users, but a user is still required. Every client/project has a generic user.
    
    #Incident Page
    driver.get(
        'https://<instance>.com/nav_to.do?uri=%2Fincident.do%3Fsys_id%3D-1%26sysparm_stack%3Dincident_list.do%3Fsysparm_query%3Dactive%3Dtrue%26sysparm_domain%3D<domain sys_id>%26sysparm_query%3Dcompany%3D' + comp_id + '^contact_type=email^caller_id=' + user_id + '^category=Workplace^subcategory=Generic^u_category3=HelpDesk'
        )

    #Switch to frame
    driver.switch_to.frame(wait(10,0,'//*[@id="gsft_main"]'))
    wait(15,0,'//*[@id="incident.close_notes"]') #Page bottom

    #Drop Downs
    wait(5,0,'//*[@id="incident.contact_type"]/option[5]').click()
    wait(5,0,'//*[@id="incident.impact"]/option[3]').click()
    wait(5,0,'//*[@id="incident.urgency"]/option[3]').click()
    driver.execute_script("g_form.setValue('category','Workplace')")
    driver.execute_script("g_form.setValue('subcategory','Generic')")
    driver.execute_script("g_form.setValue('u_category3','HelpDesk')")

    #Set Remaining Fields
    driver.execute_script("g_form.setValue('assignment_group',{})".format(json.dumps(group(queue))))
    driver.execute_script("g_form.setValue('short_description',{})".format(json.dumps(short)))
    driver.execute_script("g_form.setValue('description',{})".format(json.dumps(ticketBody)))

    #Tyme
    pyperclip.copy(datetime.strftime(m.senton, '%Y-%m-%d %H:%M:%S'))
    wait(5,0,'//*[@id="incident.u_occured"]').send_keys(Keys.CONTROL + 'v')

    #Save email, account for possible duplicate
    x= False
    while x == False:
        try:
            path= os.getcwd() + '\\Emails\\email.msg'
            msg.saveas(path)
            x= True
        except:
            try:
                num= random.randint(1,1000001)
                path= os.getcwd() + '\\Emails\\' + str(num) + 'email.msg'
                msg.saveas(path)
                x= True
            except:
                continue
            
    #Upload email, accept possible alert, delete from local dir
    wait(10,0,'//*[@id="header_add_attachment"]').click()     
    wait(10,0,'//*[@id="attachFile"]').send_keys(path)
    time.sleep(1)
    try:
        driver.switch_to.alert.accept()
        wait(15,0,'//*[@id="attachment_table_body"]/tr[2]/td/a[2]')
    except:
        pass
    wait(10,0,'//*[@id="attachment"]/div/div/header/button').click()
    os.remove(path)   

    #Check for email exclusions, insert watchlist email                                  
    if email.upper() not in exclude:
        wait(10,0,'//*[@id="incident.watch_list_unlock"]').send_keys(Keys.ENTER)
        pyperclip.copy(email)
        wait(10,0,'//*[@id="text.value.incident.watch_list"]').send_keys(Keys.CONTROL + 'v')
        wait(10,0,'//*[@id="text.value.incident.watch_list"]').send_keys(Keys.ENTER)

    #Submit and Wait
    wait(5,0,'//*[@id="6685b1c93744d7849d3b861754990ef8"]').send_keys(Keys.ENTER)
    wait(15,0,'//*[@id="sn_form_inline_stream_entries"]/ul/li')
    
    #Get ticket number
    ticketNum = wait(10,0,'//*[@id="sys_readonly.incident.number"]').get_attribute('value')     

    #Update email subject
    m.subject = re.sub('Working...','',m.subject,flags=re.I)
    m.subject = ticketNum + ' ' + re.sub('[\$\$\{\{%%].+','',m.subject)

    #Generate Template, combine into reply to preserve formats
    temp= outlook.CreateItemFromTemplate(os.getcwd()+ '\\' + 'snow_incident.msg')
    reply = m.replyall
    reply.subject = m.subject
    temp.htmlbody= temp.htmlbody.replace('ticketNummm',ticketNum)
    reply.htmlbody= temp.htmlbody + reply.htmlbody

    #Remove helpdesks from cc
    for x in reply.recipients:
        if str(x) == '<helpdesk.email>' or str(x) == '<helpdesk.email>' or str(x) == 'SAFE':
            reply.recipients.remove(x.index)

    #Send the Reply, move to completed
    reply.send
    m.save()
    m.move(destFolder)   
    
###############################################################################
#Start
###############################################################################

#Select Working Folder
print("Be prepared for a first time login (120 sec timeout).\n")
workFolder= input("Enter folder you would like to work (Default - incidents): ") or 'incidents'

#Set Folders
cwf = root_folder.Folders['Inbox'].Folders['Automated']
tasks= cwf.Folders[workFolder].Items
destFolder= root_folder.Folders['Inbox'].Folders['2019'].Folders[datetime.today().month]

#Launch driver and login handler
driver= webdriver.Chrome(executable_path=CHROMEDRIVER_PATH,options=chrome_options)

#Login
try:
    login()
except:
    #Solution to 'Successful logout' problem. Manual relaunch after erasing all driver cookies.
    driver.quit()
    os.remove(os.getcwd() + '\\Driver Data\\Default\\Cookies')
    print('Cookies cleared. Manual relaunch script and be prepared to login.')
    sys.exit()

#Get current logged in user
driver.switch_to.frame(wait(10,0,'//*[@id="gsft_main"]'))
g_user = current_snow_user()

#Run until box is empty
while len(tasks) > 0:
    #Refresh Items every 'while' pass
    tasks = cwf.Folders[workFolder].Items
    tasks.sort("ReceivedTime")
    #Iter folder, touch only unread
    for msg in tasks:
        if msg.Unread == True:
            ticket(msg)#Main
            #Refresh every 'for' pass
            tasks = cwf.Folders[workFolder].Items
            tasks.sort("ReceivedTime")


























    
