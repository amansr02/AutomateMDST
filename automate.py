import pickle
import sys
import json
import base64
import time
import os.path
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

SAMPLE_SPREADSHEET_ID = ""
SAMPLE_RANGE_NAME = ""

def initial_auth():
    username = "amansr@umich.edu"
    #check password.txt for the missing line. Could have used file io but wanted to push to github asap.
    passwd = b.decode("utf-8")

    options = Options()
    #options.headless = True
    driver = webdriver.Firefox(options=options)
    driver.get("https://accounts.google.com/signin/v2/identifier?continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&service=mail&sacu=1&rip=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin")
    assert "Gmail" in driver.title
    elem = driver.find_element_by_id("identifierId")
    elem.clear()
    elem.send_keys("amansr@umich.edu")
    elem.send_keys(Keys.RETURN)

    time.sleep(5)

    login = driver.find_element_by_id("login-visible")
    login.clear()
    login.send_keys(username)

    password = driver.find_element_by_id("password")
    password.send_keys(passwd)
    currentURL = driver.current_url
    password.send_keys(Keys.RETURN)

    time.sleep(5)
    driver.switch_to.frame('duo_iframe')
    authButton = driver.find_element_by_xpath("/html/body/div/div/div[1]/div/form/div[1]/fieldset/div[1]/button").click()
    
    while "https://mail.google.com/mail/u/0/#inbox" not in driver.current_url: 
        print("waiting...")
        pass
    print("success")
    time.sleep(6)
    return driver

def generate_email(name,standing,major):
    #return this crap
    email = """Hey """+name+""",

Hope you are well. I am Aman, a rising sophomore at MDST and I wanted to extend to you a warm welcome into the team. 

We are excited to have you on board and have many projects and questions lined up for the fall semester. We value your expertise as a """+standing+""" in """+major+""" and encourage you to be an active member of the team. 

To stay updated, join our slack channel-> https://bit.ly/2kTpR1b 

Our First Mass Meeting is in -> Thursday, September 26th from 6-7pm in SPH 1755.

We also have an upcoming Data Challenge sponsored by MIDAS => MIDAS Data Challenge. 
The registration page and info is here -> https://www.mdst.club/midas-data-challenge

Be on the lookout for emails and slack updates on project meetings for the fall semester!
Let me know if you have any questions!!

Warm Regards, 
Aman Srivastav
VP of Recruitment - MDST
"""
    return email


def selenium_interface(values): 

    driver = initial_auth()
    driver.switch_to.default_content()
    #loads gmail page
   
    for element in values:
        
        compose = driver.find_element_by_xpath("/html/body/div[7]/div[3]/div/div[2]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/div").click()
        time.sleep(2)    
        to = driver.find_element_by_name("to")
        to.send_keys(element[1])
        subject = driver.find_element_by_name("subjectbox")
        subject.send_keys("Welcome to the Michigan Data Science Team")
        body = driver.find_element_by_xpath("""//div[@aria-label="Message Body"]""")
        name = element[2].split()
        #setup conditional stuff
        if(element[4]=="Literature, Science, and the Arts"):
            element[4]= "LSA"
        if "Freshman" in element[6]:
            element[6]="Freshman"
        elif "Sophomore" in element[6]: 
            element[6]="Sophomore"
        elif "Junior" in element[6]:
            element[6]="Junior"
        elif "Senior" in element[6]: 
            element[6]="Senior"
        email = generate_email(name[0],element[6],element[4])
        body.send_keys(email)
        body.send_keys(Keys.ESCAPE)

    # submit = driver.find_element_by_xpath("""//*[@id=":nu"]""")
    # submit.click()


def sheet_auth(start,end): 

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    global SAMPLE_SPREADSHEET_ID
    global SAMPLE_RANGE_NAME
    
    

    SAMPLE_SPREADSHEET_ID = '1iIS1nezxe1ddNzAYaTdV8TN93SamoKE2WRu9qbDkT68'
    SAMPLE_RANGE_NAME = 'Form Responses 1!' + start + ':'+end

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)

    return service


def main():
    argv = sys.argv
    #replace with metadata file
    sender = "amansr@umich.edu"
    to = ""
    subject = "Welcome to MDST"
    start = "A774"
    end = "R863"    

    service = sheet_auth(start,end)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        with open("user-data.json","w") as outfile: 
            json.dump(values,outfile,indent=4)
    
    selenium_interface(values)
    
if __name__ == '__main__':
    main()


#deprecated methods
## in main: 
#gmail
    # userList = []
    # service = gmail_auth()
    # f = open("user-data.json")
    # data = json.load(f)
    # for element in data:
    #     name = element[2].split()
    #     name = name[0]
    #     userList.append({
    #         "name": name,
    #         "email": element[1],
    #         "major": element[4],
    #         "year": element[6],
    #         "experience": element[16]
    #     })
    # with open("user-data.json","w") as outfile:
    #     json.dump(userList,outfile,indent=4)
        
    # # Call the Gmail API
    # #results = service.users().labels().list(userId='me').execute()

    # user = userList[0]
    # to = user["email"]
    # message = create_message(sender,to,subject,"test")
    # draft = create_draft(service,"amansr@umich.edu",message)

# def gmail_auth(): 

#     """Shows basic usage of the Gmail API.
#     Lists the user's Gmail labels.
#     """
#     creds = None
#     # The file token.pickle stores the user's access and refresh tokens, and is
#     # created automatically when the authorization flow completes for the first
#     # time.
#     if os.path.exists('gtoken.pickle'):
#         with open('gtoken.pickle', 'rb') as token:
#             creds = pickle.load(token)
#     # If there are no (valid) credentials available, let the user log in.
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 'gcredentials.json', GSCOPES)
#             creds = flow.run_local_server()
#         # Save the credentials for the next run
#         with open('gtoken.pickle', 'wb') as token:
#             pickle.dump(creds, token)
#     service = build('gmail', 'v1', credentials=creds)
#     return service

# def create_message(sender, to, subject, message_text):
# #     """Create a message for an email.

# #   Args:
# #     sender: Email address of the sender.
# #     to: Email address of the receiver.
# #     subject: The subject of the email message.
# #     message_text: The text of the email message.

# #   Returns:
# #     An object containing a base64url encoded email object.
# #   """
#     message = MIMEText(message_text)
#     message['to'] = to
#     message['from'] = sender
#     message['subject'] = subject
#     encoded_message = base64.urlsafe_b64encode(message.as_bytes())
#     return {'raw': encoded_message.decode()}

# def create_draft(service, user_id, message_body):
# #   """Create and insert a draft email. Print the returned draft's message and id.

# #   Args:
# #     service: Authorized Gmail API service instance.
# #     user_id: User's email address. The special value "me"
# #     can be used to indicate the authenticated user.
# #     message_body: The body of the email message, including headers.

# #   Returns:
# #     Draft object, including draft id and message meta data.
# #   """
#     try:
#         message = {'message': message_body}
#         draft = service.users().drafts().create(userId=user_id, body=message).execute()
#         print('Draft id: %s\nDraft message: %s' % (draft['id'], draft['message']))
#         return draft
#     except HttpError as e:
#         print('An error occurred',e)
#         return None
