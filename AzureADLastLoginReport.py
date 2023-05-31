"""
    .NOTES
    ===========================================================================
    Created with: 	vscode
    Created on:   	5/9/2023  1:46 PM
    Created by:   	Vikas Sukhija
    Organization:
    Filename:     	cleanupguestaccounts.py
    ===========================================================================
    .DESCRIPTION
    This script will report on AzureAd last login and can send the report on designated emaail address
    also you can filter the reports based on days, like whoever has not loggedin since 90days so 
    that in case you want to deactivate guest accounts or other accounts based on that.     

    User.Read.All and AuditLog.Read.All permissions are required for the APP
"""
import configparser
import os
import csv
import requests
import json
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta
from DPAPI  import CryptUnprotectData
from vsadmin import create_file,write_log, send_mail, recycle_logs
########################log and variables####################
log_file = create_file("logs", "AzureADLastLoginReport", "log")
report_file_full = create_file("report", "ReportFull-AzureADLastLoginReport", "csv")
report_file_Actual = create_file("report", "ReportActual-AzureADLastLoginReport", "csv")
configfile = "C:\\identity\\Config.ini"
sender_address = "DoNotReplyPy@labtest.com"
rcpt_to = "Reports@labtest.com"
smtp_server = 'smtpserver'

today = datetime.today()
comparedate = today + timedelta(days=-90)
current_directory = os.getcwd()
logs_folder = current_directory + "\\" + "logs"
reports_folder = current_directory + "\\" + "report"

config = configparser.ConfigParser()
config.read(configfile)
##########################Fetch access token ##################
write_log("Start............Script", log_file)
try:
 accessstoken_file = config['o365graphtoken']['stokenfile']
 tree = ET.parse(accessstoken_file)
 root = tree.getroot()
 auth_token =  CryptUnprotectData(bytes.fromhex(root[0][1][4].text)).decode().replace('\x00', '')
except Exception as error_occured:
  write_log("exception Occured" + str(error_occured), log_file,"Error")
  send_mail(smtp_server, sender_address,rcpt_to,"Error - o365GroupNoOwnerNomembers",str(error_occured))
  quit()
##########################Start Main script####################
write_log("Start querying Graph", log_file)
coll_guestaccounts=[]
uri_endpoint = "https://graph.microsoft.com/beta/users?$filter=(userType eq 'Guest')&$select=displayName,userPrincipalName, mail, id, CreatedDateTime, signInActivity, UserType, assignedLicenses&$top=999" # end point
head =  {'Authorization': 'Bearer ' + auth_token}
response = requests.get(uri_endpoint, headers=head)
json_groups_string = json.dumps(response.json()) #striing values from json
json_groups_dict = json.loads(json_groups_string) # convert to dict
coll_guestaccounts = json_groups_dict['value'] # list
next_link = json_groups_dict['@odata.nextLink']
write_log("Initial Graph request submitted", log_file)
count = 0
rcount = 0
while True:
    try:
        count = count + 1
        rcount = rcount + 1
        write_log('Page count ' + str(count), log_file)
        if rcount == 250:
           rcount = 0
           accessstoken_file = config['o365graphtoken']['stokenfile']
           tree = ET.parse(accessstoken_file)
           root = tree.getroot()
           auth_token =  CryptUnprotectData(bytes.fromhex(root[0][1][4].text)).decode().replace('\x00', '')
           head =  {'Authorization': 'Bearer ' + auth_token}
        coll_guestaccounts_1=[]
        json_groups_string = json_groups_dict = None
        response = requests.get(next_link, headers=head)
        json_groups_string = json.dumps(response.json())
        json_groups_dict = json.loads(json_groups_string)
        coll_guestaccounts_1 = json_groups_dict['value'] # list
        next_link = json_groups_dict['@odata.nextLink']
        coll_guestaccounts.extend(coll_guestaccounts_1)
    except KeyError: 
        write_log("Key Not Found", log_file)
        break
    except Exception as error_occured:
        write_log("exception Occured" + str(error_occured), log_file,"Error")
        send_mail(smtp_server, sender_address,rcpt_to,"Error - cleanupguestaccounts",str(error_occured))

write_log('Fetched Accounts: ' + str(len(coll_guestaccounts)), log_file)
######################Extract guest accounts less than 90 days#####
date_format1 = "%Y-%m-%dT%H:%M:%SZ"
date_format2 = "%Y-%m-%d %H:%M:%S.%f"
Result_list = list(filter(lambda x: (datetime.strptime(str(x['createdDateTime']),date_format1) < datetime.strptime(str(comparedate),date_format2)),coll_guestaccounts))

##############Define class#########################################
class guestclass:
  def __init__(self, UPN, DisplayName, Email, Id, Created, LastSignInDateTime, LastNonInteractiveSignInDateTime, UserType, IsLicensed, Status):
    self.UPN = UPN
    self.DisplayName = DisplayName
    self.Email = Email
    self.Id = Id
    self.Created = Created
    self.LastSignInDateTime = LastSignInDateTime
    self.LastNonInteractiveSignInDateTime = LastNonInteractiveSignInDateTime
    self.UserType = UserType
    self.IsLicensed = IsLicensed
    self.Status = Status

coll_guest = []
for g in Result_list:
 if g['userType'] == 'Guest':
   if 'signInActivity' not in g.keys():
      if g['assignedLicenses'] ==[]:
        coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],"","",g['userType'],'False','NullData'))
      else:
        coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],"","",g['userType'],'True','NullData'))
   
   elif g['signInActivity']['lastSignInDateTime'] is None and g['signInActivity']['lastNonInteractiveSignInDateTime'] is not None:
      if datetime.strptime(str(g['signInActivity']['lastNonInteractiveSignInDateTime']),date_format1) < datetime.strptime(str(comparedate),date_format2):
        if g['assignedLicenses'] ==[]:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','Deactivate'))
        else:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','Deactivate'))
      else:
         if g['assignedLicenses'] ==[]:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','NoAction'))
         else:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','NoAction'))

   elif g['signInActivity']['lastSignInDateTime'] is not None and g['signInActivity']['lastNonInteractiveSignInDateTime'] is None:
      if datetime.strptime(str(g['signInActivity']['lastSignInDateTime']),date_format1) < datetime.strptime(str(comparedate),date_format2):
        if g['assignedLicenses'] ==[]:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','Deactivate'))
        else:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','Deactivate'))
      else:
         if g['assignedLicenses'] ==[]:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','NoAction'))
         else:
            coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','NoAction'))
   else:
      if datetime.strptime(str(g['signInActivity']['lastSignInDateTime']),date_format1) > datetime.strptime(str(g['signInActivity']['lastNonInteractiveSignInDateTime']),date_format1):
         if datetime.strptime(str(g['signInActivity']['lastSignInDateTime']),date_format1) < datetime.strptime(str(comparedate),date_format2):
            if g['assignedLicenses'] ==[]:
             coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','Deactivate'))
            else:
             coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','Deactivate'))
         else:
            if g['assignedLicenses'] ==[]:
                coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','NoAction'))
            else:
                coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','NoAction'))
      else:
         if datetime.strptime(str(g['signInActivity']['lastNonInteractiveSignInDateTime']),date_format1) < datetime.strptime(str(comparedate),date_format2):
            if g['assignedLicenses'] ==[]:
             coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','Deactivate'))
            else:
             coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','Deactivate'))
         else:
            if g['assignedLicenses'] ==[]:
                coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'False','NoAction'))
            else:
                coll_guest.append(guestclass(g['userPrincipalName'],g['displayName'],g['mail'],g['id'],g['createdDateTime'],g['signInActivity']['lastSignInDateTime'],g['signInActivity']['lastNonInteractiveSignInDateTime'],g['userType'],'True','NoAction'))

###########################export to csv###########################
with open(report_file_full, mode='w', newline='', encoding='utf-8') as csv_file:
    fieldnames = ['UPN', 'DisplayName', 'Email', 'Id', 'Created', 'LastSignInDateTime', 'LastNonInteractiveSignInDateTime', 'UserType', 'IsLicensed', 'Status']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()
    for guser in coll_guest:
        #write_log("exporting..." + str(guser.UPN), log_file)
        writer.writerow(
            {'UPN': guser.UPN, 'DisplayName': guser.DisplayName,'Email': guser.Email,'Id': guser.Id,'Created': guser.Created,'LastSignInDateTime': guser.LastSignInDateTime,'LastNonInteractiveSignInDateTime': guser.LastNonInteractiveSignInDateTime,'UserType': guser.UserType,'IsLicensed': guser.IsLicensed,'Status': guser.Status}) 

with open(report_file_Actual, mode='w', newline='', encoding='utf-8') as csv_file:
    fieldnames = ['UPN', 'DisplayName', 'Email', 'Id', 'Created', 'LastSignInDateTime', 'LastNonInteractiveSignInDateTime', 'UserType', 'IsLicensed', 'Status']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()
    for guser in coll_guest:
        if guser.Status == 'Deactivate' or guser.Status == 'NullData':
            #write_log("exporting..." + str(guser.UPN), log_file)
            writer.writerow(
                {'UPN': guser.UPN, 'DisplayName': guser.DisplayName,'Email': guser.Email,'Id': guser.Id,'Created': guser.Created,'LastSignInDateTime': guser.LastSignInDateTime,'LastNonInteractiveSignInDateTime': guser.LastNonInteractiveSignInDateTime,'UserType': guser.UserType,'IsLicensed': guser.IsLicensed,'Status': guser.Status}) 
############################RecycleLogs############################
recycle_logs(logs_folder,60)
recycle_logs(reports_folder,60)