#General Notes
"""
This script is dependent on a lightweight database (shelve in this example) to perform its lookups and guesses.

I utilized Pandas to parse ServiceNow data and  generate many of the dictionaries and lists stored in the shelve object.

Was not designed for use by anyone beyond those utilizing my other email processing scripts and myself.
"""
#Problem
"""
Email processing scripts relied on one or two persons to format the message's subject line with uniquely encased strings.

Ex: {{assignment queue}} %%client%%. These two values are required for each email.

Therefore, the inbox's performance was dependent on the person formatting the subject lines to instinctively know those
two values for every email.

Not a problem for a regular or experienced individuals responsible for the box but I saw a need for a utility for when
our email expert wasn't going to make it in. Additionally, it acts a quality-of-life tool for anyone monitoring the inbox,
experienced or otherwise.
"""

#Solution
"""
A script that parses a folder and makes a best guess as to what the email's {{client}} and %%assignment group%% most likely is based of total history.

Individual Email Flow:

Note: If a client is found, it will not change unless a queue is selected that is directly related to a particular client. 

1. Look in 'people_client' for user's email and associated client. Look in body for queue name, if known queue associated with a client is found (deskside) select that client instead. Done.
2. If no queue found, look in 'people_queue' for user and guess index 0 of that person's queue list (which is ordered by number of tickets to that queue by that person). If deskside queue, select that client instead. Done.
3. If queue was guessed, move to respective guess folder.
4. User checks the guessed queue for accuracy. If wrong, user leaves the email and next script run, the next queue in the persons queue list will be guessed. If deskside then select that client instead.
5. If no queue is accurately determined, move to 'Null Queue' and mark unread so user realizes the folder is not empty.
"""
import win32com.client
import os, time
import re
import shutil
import csv
import shelve
from pathlib import Path
from fuzzywuzzy import fuzz

############################################################
#Functions
############################################################

#Parse email body for ServiceNow Assignment Group-like string
def regex_queue(msg):
    pat = '\S+\.cndt'
    try:
        temp = re.findall(pat,msg.body,flags=re.I)[0].lower()
        if temp.lower() not in lqlist:
            cur = 0
            for i in lqlist:
                if fuzz.partial_ratio(temp,i) > 85 and fuzz.partial_ratio(temp,i) > cur:
                    q = i
                    cur = fuzz.partial_ratio(temp,i)
        else:
            q = temp
        return q.lower()
    except:
        return None

#Select passed index of person's queue list
def guess_queue(ind):
    try:
        return people_queue[email][ind].lower()
    except:
        return None

#Look up person's associated project
def lookup_client(msg):
    if email in people_client:
        return people_client[email]
    else:
        return None

############################################################
#Constants
############################################################


emaildb = Path(r"Data\emaildb")
data_path = Path(r'\\Path\to\my\machine\Data')

#If 'Data' directory does not exist, then copy from my machine
if not os.path.exists('Data'):
    shutil.copytree(data_path,'Data')
time.sleep(3)

#Unpack Database
with shelve.open(emaildb.as_posix()) as db:
    lqlist = db['lowercase_queues_list']
    desks = db['deskqueue_client_dict']
    people_client = db['people_client_dict']
    people_queue = db['people_qlist']
    location_client = db['location_client']
    cost_client = db['cost_client']


#Get Outlook API/Namespace
namespace = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')

#Folders
prime = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Main'].Items
guessers = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Guess'].Items
guess_folder = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Guess']
queue_found = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Queue Found']
null_found = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Null Queue']

#Guess folder first, touch only unread mails
for m in guessers:
    if m.unread == True:

        #Remove existing queue
        m.subject = re.sub('{{.+}}','',m.subject)

        #Extract email address
        email= m.senderemailaddress.lower()
        if '@' not in email:
            email= m.sender.GetExchangeUser().PrimarySmtpAddress.lower()

        #Adjust Index and select new
        index = int(re.findall('(\d)$',m.subject)[0]) + 1
        queue = guess_queue(index)

        #Edit and save or mark read
        if queue != None:
            if '%%' not in m.subject and queue in desks:
                client = desks[queue]
                m.subject = m.subject + ' %%' + client.upper() + '%%' + ' {{' + queue.upper() + '}} ' + str(index)
                m.save()
            else:
                m.subject = m.subject + ' {{' + queue.upper() + '}} ' + str(index)
                m.save()
        else:
            m.unread = False

#Main folder, touch only unread
while len(prime) > 0:
    for m in prime:
        if m.unread == True:
            #############################################
            #Verify there's an email address to work with
            #############################################
            
            #Try to find email
            try:
                email= m.senderemailaddress.lower()
                if '@' not in email:
                    email= m.sender.GetExchangeUser().PrimarySmtpAddress.lower()
            except:
                #Try to find queue
                queue = regex_queue(m)
                if queue in desks: #Deskside check
                    client = desks[queue]
                    m.subject = m.subject + ' %%' + client.upper() + '%%'
                    m.save()
                if queue != None: #Save found queue
                    m.subject = m.subject + ' {{' + queue.upper() + '}}'
                    m.save()
                    m.move(queue_found)
                else: #Couldn't determine anything
                    m.move(null_found)

                prime = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Main'].Items
                continue

            ############################
            #An Email Address was found
            ############################

            #Init vars
            client = None
            queue = None
            found = False

            #Look up client and save
            client = lookup_client(m)
            if client != None:
                m.subject = m.subject + ' %%' + client.upper() + '%%'
                m.save()

            #Search for queue, guess if not
            queue = regex_queue(m)
            if queue == None:
                queue = guess_queue(0)
            else:
                found = True

            #Update subject line else mark read
            if queue != None:
                if queue in desks and '%%' not in m.subject:#If not already determined, check for deskside
                    client = desks[queue]
                    m.subject = m.subject + ' %%' + client.upper() + '%%' + ' {{' + queue.upper() + '}} 0'
                    m.save()
                else:
                    m.subject = m.subject + ' {{' + queue.upper() + '}} 0'
                    m.save()
            else:
                m.unread = False

            #Move based on guess or not, move to null if nothing else
            if m.unread == True:
                if found == True:
                    m.subject = re.sub('(\s\d)$','',m.subject)
                    m.save()
                    m.move(queue_found)
                else:
                    m.move(guess_folder)
            else:
                m.move(null_found)

            #Refresh 'Main'
            prime = namespace.Folders['my.helpdesk@organization.net'].Folders['Inbox'].Folders['Automated'].Folders['Auto Format'].Folders['Main'].Items

#Mark emails in null unread so not forgotten
for m in null_found.Items:
    m.unread = True
                
        







































