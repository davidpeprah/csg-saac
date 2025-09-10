#
#
#  Author: David Peprah
#
#

from __future__ import print_function
import pickle, pdb, re
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from configparser import ConfigParser
import subprocess, sys, logging, argparse, platform
from jinja2 import Environment, FileSystemLoader
from datetime import datetime

# Custom functions from lib folder
from lib.checkUserGS import checkUser
from lib.sendEmail import sendMessage, CreateMessageWithAttachment

# set up the Jinja2 environment to load templates from the 'templates' directory
env = Environment(loader=FileSystemLoader('templates'))

# Check log file and create it if it does not exist
def check_log_file(filepath: str):
    if not os.path.exists(filepath):
        os.makedirs(os.path.dirname(filepath))
        with open(fr"{filepath}", "w") as file: 
            pass


def readSheet(response_sheet):
    global sheet
    global adminAlerts

    RANGE_NAME = response_sheet + Range

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        logging.warning(f"No data found in the range {RANGE_NAME} of the spreadsheet {SPREADSHEET_ID}.")
        return None 
    
    logging.info(f"Data retrieved successfully from the range {RANGE_NAME} of the spreadsheet.")
    row_count = 2
    for row in values:
        logging.debug(f"Processing row {row_count}: {row}")
        try:
            fname = row[1].title().strip()
            mname = row[2].title().strip()
            lname = row[3].title().strip()
            pEmail = row[4].strip()
            department = row[5].lower().strip()
            jobTitle = row[6].title().strip()
            jobRole = row[13].lower().strip()
            googleSharedDrive = row[7].strip()
            localSharedDrive = row[8].strip()
            phoneExtension = row[9].strip()
            otherInfo = row[10].strip()
            curEmpEmail = row[11].lower().strip()
            entryType = row[12].strip()
            ops_status = row[15].strip() if len(row) > 15 else ""


            pwshell_testing = "false"
            if testing:
                logging.debug("Running in testing mode, no changes will be made to the AD.")
                curEmpEmail = admin
                adminAlerts = admin
                pwshell_testing = "true"


            if entryType == "NEW":
                logging.info(f"Creating account for new hire: {fname} {lname}, Job Role: {jobRole}, Department: {department}")
                
                if curEmpEmail.split("@")[0] not in authUsers:
                    logging.warning(f"Unauthorized user {curEmpEmail} trying to create a district email account.")
                    # Update the status and entry type to Denied and OLD respectively
                    UpdateStatus(response_sheet,row_count,status_msg("3"),"Access Denied")
                    UpdateEntryType(response_sheet,row_count,"OLD")
                    
                    # Send an email to the unauthorized user and copy the admin alerts
                    send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year}, recipient=curEmpEmail, subject="Unauthorized User Attempt",
                                            template_name="access_denied.html", with_attachment=False,cc=adminAlerts)
                    
                    # Skip to the next row
                    row_count += 1
                    continue

                
                #Get AD groups
                jobrole_ = "".join(re.split('[^a-zA-Z0-9]+', jobRole.lower()))
                department_ = "".join(re.split('[^a-zA-Z0-9]+', department.lower()))
                adgrps = ADGroups(jobrole_, department_)
                adorganizationalunit = organizationalUnit(jobrole_, department_)
                
                
                # check if organizational unit is found
                if adorganizationalunit is None:
                    logging.error(f"Organizational Unit not found for Job Role: {jobRole}, Department: {department}. Cannot create account for {fname} {lname}.")
                    UpdateStatus(response_sheet,row_count,status_msg("2"),"Error: Organizational Unit not found")
                    # Send an email to ERC Team when there is an error
                    send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year, "new_hire_email": "N/A", "error_message": f"Organizational Unit not found for Job Role: {jobRole}, Department: {department}. Cannot create account for {fname} {lname}.", 
                                                  "new_hire_jrole": jobRole, "new_hire_dpart": department, "new_hire_adgroups": adgrps, "new_hire_ou": "N/A"}, 
                                                recipient=admin, subject="Account Creation Error",template_name="account_creation_error.html", with_attachment=False,cc=adminAlerts)
                    row_count += 1
                    continue
                
                # Send Data to Powershell
                logging.debug(f"Creating account for {fname} {lname} AD groups: {adgrps}, Job Role: {jobRole}, Department: {department}, OU: {adorganizationalunit}")
                 # Call Powershell script to create account
                createAcc = subprocess.Popen(["Powershell.exe", "-File", r"lib\setAcc.ps1",
                                              "-FirstName", fname,
                                              "-MiddleName", mname,
                                              "-LastName", lname,
                                              "-jobrole", jobRole.lower(),
                                              "-department", department.lower(),
                                              "-adgroups", adgrps,
                                              "-oupath", adorganizationalunit,
                                              "-jobtitle", jobTitle,
                                              "-testing", pwshell_testing], stdout=subprocess.PIPE)

                # Read Information from Powershell
                response = str(createAcc.communicate()[0][:-2], 'utf-8')
                logging.debug(response)
                status = ""; email = ""; update = ""; other_output = ""
                res = response.split('\r\n')
                status, email, update = res[-3], res[-2], res[-1]
                if len(res) > 3:
                    other_output = ";".join(res[:-4])
                    logging.debug(f"Additional output from PowerShell script: {other_output}")

                if (status == "1"):
                    logging.info(f"Account for {fname} {lname} created successfully: {email}")
                    UpdateStatus(response_sheet,row_count,status_msg(status),update)
                    UpdateMail(response_sheet,row_count,email)
                    UpdateEntryType(response_sheet,row_count,update)

                    
                    # Send email notification other information to the IT department to set up laptop, phone extension etc
                    logging.debug(f"Sending new hire IT request email to {adminAlerts} for {fname} {lname}")
                    send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year, "new_hire_email": email, 
                                                  "new_hire_jrole": jobRole.title(), "new_hire_dpart": department.title(), "laptop_preference": laptop_preference, "googleSharedDrive": googleSharedDrive, 
                                                  "localSharedDrive": localSharedDrive, "needPhoneExtension": needPhoneExtension, "phoneExtension": phoneExtension, 
                                                  "otherInfo": otherInfo}, 
                                                recipient=adminAlerts, subject="New Employee IT Request",template_name="employee_requests.html", with_attachment=False,cc=curEmpEmail)
                    

                elif (status == "2"):
                    UpdateStatus(response_sheet,row_count,status_msg(status),update)
                    UpdateEntryType(response_sheet,row_count,update)
                    logging.error(f"Error creating account for {fname} {lname}: {update}")  
                    # Send an email to ERC Team when there is an error
                    send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year, "new_hire_email": email, "error_message": update + f"<br> Additional Info: {other_output}", 
                                                  "new_hire_jrole": jobRole, "new_hire_dpart": department, "new_hire_adgroups": adgrps, "new_hire_ou": adorganizationalunit}, 
                                                recipient=admin, subject="Account Creation Error",template_name="account_creation_error.html", with_attachment=False,cc=adminAlerts)
                                                                                       
            elif ops_status == "Pending":
                logging.info(f"Processing pending account for: {fname} {lname}, Job Role: {jobRole}, Department: {department}")
                
                personal_email = str(pEmail)
                firstName = str(fname)
                lastName = str(lname)
                EmpEmail = str(curEmpEmail)
                newemail = row[14]

                try:
                    check = checkUser(newemail, dir_nav)
                    
                    if check.lower() == newemail.lower():
                        logging.info(f"Account for {fname} {lname} verified in G-Suite: {newemail}")
                        UpdateStatus(response_sheet,row_count,status_msg("0"),"Account Successfully confirm in G-Suite")
                        passw = password()
                        logging.debug(f"Resetting password for {newemail}")

                        # Call Powershell script to reset password
                        passReset = subprocess.Popen(["Powershell.exe", "-File", "lib\\resetPass.ps1",
                                                                                "-Email", newemail,
                                                                                "-NewPassword", passw,
                                                                                "-testing", pwshell_testing 
                                                                                ], stdout=subprocess.PIPE)
                        response = str(passReset.communicate()[0][:-2], 'utf-8')
                        res = response.split('\r\n')
                        status = ""; update = ""; other_output = ""
                        status, update = res[-2], res[-1]
                        if len(res) > 2:
                            other_output = ";".join(res[:-3])
                            logging.debug(f"Additional output from PowerShell script: {other_output}")
                        

                        

                        #Send account information to the new staff personal email and notify the employee who made the entry
                        if (status == "0"):
                            logging.info(f"Password reset successful: {update}")
                            
                            if (personal_email):
                                logging.info(f"Sending account information to new hire personal email: {personal_email}")
                                # Notify the new hire of their account information
                                send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year,
                                                            "username": f"CSG\\{newemail.split('@')[0]}", "email": newemail, "password": passw}, 
                                                    recipient=personal_email, subject="Account Information from CSG",template_name="new_hire_account_notification.html", with_attachment=False)
                     
                                # Notify the employee who made the entry
                                logging.info(f"Notification email sent to {EmpEmail} about account creation for {firstName} {lastName}.")   
                                send_email_notification(data={"current_year": datetime.now().year,"employee_name": f"{fname} {lname}", "new_hire_jrole": jobRole, "new_hire_dpart": department,
                                                            "personal_email": personal_email}, 
                                                        recipient=curEmpEmail, subject=f"Account for {firstName} {lastName}  Completed Successfully",template_name="personal_email.html", with_attachment=False,cc=adminAlerts)
                            
                                 
                            else:
                                logging.warning(f"No personal email provided for new hire {firstName} {lastName}. Cannot send account information.")
                                # Send an email to ERC Team when there is an error
                                send_email_notification(data={"current_year": datetime.now().year,"username": f"CSG\\{newemail.split('@')[0]}", "email": newemail, "password": passw,
                                                            "employee_name": f"{firstName} {lastName}", "new_hire_jrole": jobRole, "new_hire_dpart": department}, 
                                                    recipient=curEmpEmail, subject=f"Account for {firstName} {lastName}  Completed Successfully",template_name="no_personal_email.html", with_attachment=False,cc=adminAlerts)
                        

                        elif (status == "1"):
                            logging.error(f"Password reset failed: {update}")
                            #Send email to development team
                            send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year, "new_hire_email": newemail, "error_message": update, 
                                                  "new_hire_jrole": jobRole, "new_hire_dpart": department, "new_hire_adgroups": adgrps, "new_hire_ou": adorganizationalunit}, 
                                                recipient=admin, subject="New Account Password Reset Error",template_name="account_creation_error.html", with_attachment=False,cc=adminAlerts)

                            
                            # Notify the new hire of their account information
                            send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year,
                                                        "username": f"CSG\\{newemail.split('@')[0]}", "email": newemail}, 
                                                    recipient=adminAlerts, subject="New Account Password Reset Error",template_name="password_reset_failed.html", with_attachment=False, cc=curEmpEmail)
                    else:
                        logging.warning(f"Account for {fname} {lname} not found in G-Suite yet: {newemail}") 
                            
                       
                except HttpError as err:
                    if err.resp.status in [404,]:
                        UpdateStatus(response_sheet,row_count,status_msg("1"),"Account has not been created in Google Console yet")
                        logging.warning(f"Account for {fname} {lname} not found in G-Suite yet: {newemail}")
                    else:
                        UpdateStatus(response_sheet,row_count,status_msg("2"),"Error occured when trying to verify account in Google console")
                        logging.error(f"Error checking account for {fname} {lname} in G-Suite: {err}")

                except Exception as e:
                    UpdateStatus(response_sheet,row_count,status_msg("2"),"Error occured when trying to verify account in Google console")
                    logging.error(f"Unexpected error checking account for {fname} {lname} in G-Suite: {e}")
                    # Send an email to ERC Team when there is an error
                    send_email_notification(data={"new_hire_fname": fname, "new_hire_lname": lname, "current_year": datetime.now().year, "new_hire_email": newemail, "error_message": str(e), 
                                                  "new_hire_jrole": jobRole, "new_hire_dpart": department, "new_hire_adgroups": adgrps, "new_hire_ou": adorganizationalunit}, 
                                                recipient=admin, subject="Account Creation Error",template_name="account_creation_error.html", with_attachment=False,cc=adminAlerts)

            else:
                logging.debug(f"Row {row_count} does not match any known entry types. Skipping...")


        except IndexError as e:
            logging.error(f"Row {row_count} is missing some data, error_msg: {str(e)}. Skipping to the next row.")
            
            send_email_notification(data={"new_hire_fname": "not applicable", "new_hire_lname": "not applicable", "current_year": datetime.now().year, "new_hire_email": "not applicable", "error_message": f"{str(e)} <br> Missing data in the row {row_count}, please check the entry.", 
                                                  "new_hire_jrole": "not applicable", "new_hire_dpart": "not applicable", "new_hire_adgroups": "not applicable", "new_hire_ou": "not applicable"}, 
                                                recipient=admin, subject="Account Creation Error",template_name="account_creation_error.html", with_attachment=False,cc=adminAlerts)
            row_count += 1
            continue
        except Exception as e:
            logging.error(f"Unexpected error processing row {row_count}, error_msg: {str(e)}. Skipping to the next row.")
            
            send_email_notification(data={"new_hire_fname": "not applicable", "new_hire_lname": "not applicable", "current_year": datetime.now().year, "new_hire_email": "not applicable", "error_message": f"{str(e)} <br> Check row {row_count} entry.", 
                                                  "new_hire_jrole": "not applicable", "new_hire_dpart": "not applicable", "new_hire_adgroups": "not applicable", "new_hire_ou": "not applicable"}, 
                                                recipient=admin, subject="Account Creation Error",template_name="account_creation_error.html", with_attachment=False,cc=adminAlerts)
            row_count += 1
            continue
        
        row_count += 1


def UpdateStatus(sh,address,status,update):
    global sheet
    logging.debug(f"Updating status for row {address} in sheet {sh} to '{status}' with comment '{update}'") 
    RANGE_NAME = sh + "!" + StatusColAdd + str(address) + ":" + CommentColAdd + str(address)

    body_value = {"majorDimension": "ROWS", "values": [[str(status),str(update)]]}

    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()

    
def UpdateMail(sh,address,message):
    global sheet
    logging.debug(f"Updating email for row {address} in sheet {sh} to '{message}'")
    RANGE_NAME = sh + "!" + NewMailColAdd + str(address)

    body_value = {"majorDimension": "ROWS", "values": [[str(message)]]}


    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()


def UpdateEntryType(sh,row_add,message):
    global sheet
    logging.debug(f"Updating entry type for row {row_add} in sheet {sh} to '{message}'")
    RANGE_NAME = sh + "!" + EntryTypeColAdd + str(row_add)

    body_value = {"majorDimension": "ROWS", "values": [[str("OLD")]]}

    
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()
    
    
def status_msg(status):
    status_dict = {
        "0": "Done",
        "1": "Pending",
        "2": "Error",
        "3": "Denied"
    }
    return status_dict[status]
        

def password():
    import random
    import array
    password = " "

    def _generate_random_password(length=14):
        # Generate a random password of specified length
        characters = array.array('u', 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+')
        return ''.join(random.choice(characters) for _ in range(length))
    

    if os.path.exists(f'config{dir_nav}dictionary'):
        # If dictionary file exists, read from it
        with open(f'config{dir_nav}dictionary', 'r', encoding='utf-16') as f:
            words = f.read().splitlines()

        if words:
            # Remove duplicate words and filter out empty strings
            words = list(set(word.strip() for word in words if word.strip())) 
        
            if len(words) > 30000:
                # Generate a password using 3 random words from the dictionary
                # Ensure the password is at least 14 characters long
                while len(password) < 14:
                    random.shuffle(words)  # Shuffle the list of words
                    random_words = random.sample(words, 3)  # Get 3 random words
                    for i in range(len(random_words)):
                        random_words[i] = random_words[i].title() # Capitalize first letter of each word
                    password = "-".join(random_words)  # Join them with a hyphen

            else:
                # If dictionary has less than 20000 words, generate a random password
                password = _generate_random_password()
        else:
            # If dictionary has less than 20000 words, generate a random password
            password = _generate_random_password()
    else:
        # If dictionary file does not exist, generate a random password
        password = _generate_random_password()

    # Ensure the password is at least 14 characters long
    return(password)

    
def ADGroups(jobrole, department):

    defaultgroups = config.get('DefaultGroups', 'defaultGroups') if (config.get('DefaultGroups', 'defaultGroups')) else ''
    if defaultgroups: defaultgroups = [grp.strip() for grp in defaultgroups.split(',')]

    grpsbyJobRole = groupsbyJobRole(jobrole)
    if grpsbyJobRole: grpsbyJobRole = [grp.strip() for grp in grpsbyJobRole.split(',')]

    grpsbyDepartment = groupsbyDepartment(department)
    if grpsbyDepartment: grpsbyDepartment = [grp.strip() for grp in grpsbyDepartment.split(',')]

    logging.debug(f"Default Groups: {defaultgroups}, Groups by Job Role: {grpsbyJobRole}, Groups by Department: {grpsbyDepartment}")
    return ",".join(defaultgroups+grpsbyJobRole+grpsbyDepartment)


def groupsbyJobRole(jobrole):
    logging.debug(f"Looking up groups for job role: {jobrole}")
    if config.options('GroupsbyJobRole'):
        if jobrole.lower() in config.options('GroupsbyJobRole'):
            logging.debug(f"Found groups for job role: {jobrole}")
            return config.get('GroupsbyJobRole', jobrole)
    logging.debug(f"No groups found for job role: {jobrole}")
    return ''


def groupsbyDepartment(department):
    logging.debug(f"Looking up groups for department: {department}")
    if config.options('GroupsbyDepartment'):
        if department.lower() in config.options('GroupsbyDepartment'):
            logging.debug(f"Found groups for department: {department}")
            return config.get('GroupsbyDepartment', department)
    logging.debug(f"No groups found for department: {department}")
    return ''

def organizationalUnit(jobrole, department):
    departmentJobrole = f"{department.lower()}{jobrole.lower()}"
    logging.debug(f"Looking up OU for combined department and job role: {departmentJobrole}")
    if config.options('OrganizationalUnits'):
        if departmentJobrole in config.options('OrganizationalUnits'):
            logging.debug(f"Found OU for combined department and job role: {departmentJobrole}")
            return config.get('OrganizationalUnits', departmentJobrole)
        elif jobrole.lower() in config.options('OrganizationalUnits'):
            logging.debug(f"Found OU for job role: {jobrole}")
            return config.get('OrganizationalUnits', jobrole)
    logging.debug(f"No OU found for job role: {jobrole} or combined department and job role: {departmentJobrole}")  
    return None

def send_email_notification(data: dict = None, recipient: str = None, subject: str = " ", file_path: str = None, file_name: str = None, template_name: str = None, with_attachment: bool = False, message: str = "TESTING EMAIL NOTIFICATION", cc: str = None):
    
    if recipient:

        send_email_message = None
        
        email_template = env.get_template(template_name)
        logging.debug(f"Using email template: {template_name}")
        rendered_email = email_template.render(data)

        if with_attachment:

            if file_path and file_name and template_name:
               
               

                # Function to send email notification
                logging.info(f"Sending email notification with attachment subject: {subject} ...")
                logging.debug(f"Email subject: {subject}")
                logging.debug(f"Email recipient: {recipient}")
                try:
                    send_email_message = sendMessage('me', CreateMessageWithAttachment(srvAccEmail, recipient, 
                                                                                subject, rendered_email, file_dir=file_path, filename=file_name, cc=cc), dir_nav)
                  
                except Exception as e:
                    logging.exception(f"Failed to send email notification with attachment: {e}")
            
            else:
                logging.critical("File path, file name or template name not provided for email notification with attachment")   
        else:
            if not template_name:
                logging.critical("Email template name not provided for email notification without attachment")
                return
            
           
            # Function to send email notification
            logging.info(f"Sending email notification without attachment subject: {subject} ...")
            logging.debug(f"Email subject: {subject}")
            logging.debug(f"Email recipient: {recipient}")
            try:
                send_email_message = sendMessage('me', CreateMessageWithAttachment(srvAccEmail, recipient, 
                                                                                    subject, rendered_email, cc=cc), dir_nav)
            except Exception as e:
                logging.exception(f"Failed to send email notification: {e}")

        # Check if the email was sent successfully
        if send_email_message:
            if send_email_message[0] == 'success':
                logging.info(f"Email notification sent to with subject {subject}")
            else:
                logging.error(f"Failed to send email notification: {send_email_message[1]}")

    else:   
        logging.critical("No recipient email provided or email template, skipping email notification")
 
def load_credentials_access():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
     # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/admin.directory.user.readonly', 'https://www.googleapis.com/auth/gmail.send']

    global sheet

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(f'config{dir_nav}token.pickle'):
        logging.debug(f"Loading credentials from config{dir_nav}token.pickle")
        with open(f'config{dir_nav}token.pickle', 'rb') as token:
            logging.debug("Credentials file found, loading credentials...")
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logging.debug("Refreshing expired credentials...")
            # If the credentials are expired or does not exist refresh them
            creds.refresh(Request())
        else:
            logging.debug("No valid credentials found, initiating authorization flow...")
            # If there are no valid credentials available, let the user log in.
            flow = InstalledAppFlow.from_client_secrets_file(f'config{dir_nav}app_auth.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(f'config{dir_nav}token.pickle', 'wb') as token:
            logging.debug(f"Saving credentials to config{dir_nav}token.pickle")
            # Save the credentials as a pickle file
            pickle.dump(creds, token)
    # Build the service object for the Sheets API
    try:
        logging.debug("Building the Sheets API service...")
        service = build('sheets', 'v4', credentials=creds)
    except HttpError as err:
        logging.critical(f"An error occurred while building the Sheets API service: {err}")
        sys.exit(1)
    logging.info("Sheets API service built successfully.")
    
    # Call the Sheets API
    sheet = service.spreadsheets()


def main():    # Load credentials and access the Google Sheets API
    global response_sheet
    load_credentials_access()
    logging.info("Google Sheets API credentials loaded and access granted.")    

     # Read the response sheet
    if testing:
        logging.warning("Running in testing mode, no changes will be made to the spreadsheet or accounts.")
        response_sheet = 'Testing Form Responses 1'

    readSheet(response_sheet)


if __name__ == '__main__':
    
    # define the directory navigation character based on the OS
    dir_nav = "\\" if platform.system() == 'Windows' else "//"
    
    # Load Config file
    config = ConfigParser()
    config.read(f'config{dir_nav}config.ini')
    if not config.sections():
        logging.critical("Configuration file is missing, empty or invalid. Exiting...")
        sys.exit(1)
    
    # Set default logging configuration
    logLevel = config.get('logs', 'logLevel' , fallback='INFO')
    logFile = config.get('logs', 'logFile', fallback=f'logs{dir_nav}csg-saac.log')
     
     # check if the log file path is valid for the current OS 
    if dir_nav not in logFile:
        logFile = logFile.replace("\\", dir_nav)
        
    check_log_file(logFile)
    
    parser = argparse.ArgumentParser(prog='CSG-SAAC',
                                     description='New Staff Onboarding and Account Creation Script',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-lL', '--logLevel', default=None, type=str, help='Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)')
    parser.add_argument('-v', '--version', action='version', version='%(prog)s 1.0.1', help='Show the program version')
    parser.add_argument('-t', '--testing', action='store_true', help='For testing purposes only, do not use in production')

    args = parser.parse_args()


    logLevel = args.logLevel.upper() if args.logLevel else logLevel.upper()
    testing = args.testing

   

    # Set up logging configuration
    numeric_level = getattr(logging, logLevel)
    if not isinstance(numeric_level, int):
        logging.CRITICAL(f"Invalid log level: {logLevel}")
        sys.exit(1)

    logging.basicConfig(level=numeric_level, 
                        format="{asctime} {levelname} {lineno}: {message}", 
                        style="{",
                        datefmt="%Y-%m-%d %H:%M:%S",
                        handlers=[ logging.FileHandler(filename=logFile, mode='a+', encoding='utf-8'),
                                    logging.StreamHandler()
                                ]
                )
        

    if 'Document' not in config.sections() or 'admin' not in config.sections():
       logging.critical("Configuration file is missing required sections: 'Document' or 'admin'")
       sys.exit(1)
        

    # The ID and range of a sample spreadsheet.
    if (config.get('Document', 'SpreadSheetID')):
        SPREADSHEET_ID = config.get('Document', 'SpreadSheetID')
        response_sheet = 'Form Responses 1' if not (config.get('Document', 'Sheets')) else config.get('Document', 'Sheets')
        Range = '!A2:R' if not (config.get('Document', 'SheetRange')) else config.get('Document', 'SheetRange')
        EntryTypeColAdd = config.get('Document', 'EntryTypeColAdd')
        NewMailColAdd = config.get('Document', 'NewMailColAdd')
        StatusColAdd = config.get('Document', 'StatusColAdd')
        CommentColAdd = config.get('Document', 'CommentColAdd')
        logging.debug(f"Spreadsheet ID: {SPREADSHEET_ID}, Sheet: {response_sheet}, Range: {Range}, EntryTypeColAdd: {EntryTypeColAdd}, NewMailColAdd: {NewMailColAdd}, StatusColAdd: {StatusColAdd}, CommentColAdd: {CommentColAdd}")
        logging.info("Spreadsheet information loaded successfully.")
    else:
        logging.critical("Spreadsheet ID is missing in the configuration file.")
        sys.exit(1)

    # Get authorized users and domains
    if (config.get('admin', 'AuthorizeUsers')):
        authUsers = list(config.get('admin', 'AuthorizeUsers', fallback='dpeprah').split(','))
        domain = config.get('admin', 'Domain')
        admin = config.get('admin', 'sysadmin',fallback='dpeprah@vartek.com')
        openticket = config.get('admin', 'openticket')
        srvAccEmail = config.get('admin', 'serviceAccEmail')
        if not srvAccEmail:
            logging.critical("Service account email is missing in the configuration file.")
            sys.exit(1)
        adminAlerts = config.get('admin', 'adminAlerts', fallback=admin)
        adminAlerts = [f"{email}@{domain}" if "@" not in email else email for email in adminAlerts.split(',')]
        adminAlerts = ",".join(adminAlerts) # Convert list back to comma-separated string
        logging.debug(f"Authorized Users: {authUsers}, Domain: {domain}, Admin: {admin}, Open Ticket: {openticket}, Service Account Email: {srvAccEmail}, Admin Alerts: {adminAlerts}")
        logging.info("Authorized users and domain loaded successfully.")
    else:
        logging.critical("Authorized users or domain is missing in the configuration file.")
        sys.exit(1)
           

    sheet = ""
    
    # call the main function 
    main()

   