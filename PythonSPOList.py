"""
    .NOTES
    ===========================================================================
    Created with: 	vscode
    Created on:   	7/14/2023  1:46 PM
    Created by:   	Vikas Sukhija
    Organization:
    Filename:     	PythonSPOList.py
    ===========================================================================
    .DESCRIPTION
    This script will read the SPO list using Graph API.
    APP registration with Delegated permissions required.
    AllSites.Read dleagted permissions
"""
import os
import requests
import json
from DPAPI  import CryptUnprotectData, CryptProtectData
#################Create config folder########################
current_directory = os.getcwd()
config_folder_path = os.path.join(current_directory, "config")
try:
    os.mkdir(config_folder_path)  # Create the new folder
    print(f"Folder  created successfully!")
except FileExistsError:
    print(f"Folder already exists.")
except Exception as e:
    print(f"An error occurred: {str(e)}")
########################log and variables####################
enc_clientid_path = current_directory + "\\" + "config" + "\\" + "encclientid.txt"
enc_Tenantid_path = current_directory + "\\" + "config" + "\\" + "enctenantid.txt"
enc_pwd_path = current_directory + "\\" + "config" + "\\" + "encpassword.txt"
enc_secret_path = current_directory + "\\" + "config" + "\\" + "encsecret.txt"
RedirectUri = "https://localhost"
tenant = "techwizard.onmicrosoft.com"
siteName = "CGQ/APIQ"
listid = "D7Z4732F-5B3B-4LA3-926N-0L179FFA5A8B"
#####################Graphurl for list#######################
guri = "https://graph.microsoft.com/v1.0/sites/techwizard.sharepoint.com:/sites/" + siteName + ":/lists/" + listid + "/items?&expand=fields&top=5000"
#####################check if password and secret exists######
if os.path.isfile(enc_clientid_path):
    print("File exists. Continuing with program execution.")
    with open(enc_clientid_path, "rb") as file:
        encrypted_clientid = file.read()
else:
    clientid = input("File does not exist. Please enter the Clientid: ")
    # Here, you can perform additional logic based on the entered Clientid
    print(f"Entered Clientid: {clientid}")
    print(f"Encrypting the enterted Clientid and saving it to {enc_clientid_path}")
    encrypted_clientid = CryptProtectData(clientid.encode())
    with open(enc_clientid_path, "wb") as file:
        file.write(encrypted_clientid)

if os.path.isfile(enc_pwd_path):
    print("File exists. Continuing with program execution.")
    with open(enc_pwd_path, "rb") as file:
        encrypted_password = file.read()
else:
    password = input("File does not exist. Please enter the password: ")
    # Here, you can perform additional logic based on the entered password
    print(f"Entered password: {password}")
    print(f"Encrypting the enterted password and saving it to {enc_pwd_path}")
    encrypted_password = CryptProtectData(password.encode())
    with open(enc_pwd_path, "wb") as file:
        file.write(encrypted_password)

if os.path.isfile(enc_secret_path):
    print("File exists. Continuing with program execution.")
    with open(enc_secret_path, "rb") as file:
        encrypted_Secret = file.read()
else:
    Secret = input("File does not exist. Please enter the Client Secret: ")
    # Here, you can perform additional logic based on the entered Client Secret
    print(f"Entered Secret: {Secret}")
    print(f"Encrypting the enterted Secret and saving it to {enc_secret_path}")
    encrypted_Secret = CryptProtectData(Secret.encode())
    with open(enc_secret_path, "wb") as file:
        file.write(encrypted_Secret)
#######################Decode passwords to be used############
decrypt_clientid = (CryptUnprotectData(encrypted_clientid)).decode()
decrypt_Tenantid = (CryptUnprotectData(encrypted_password)).decode()
decrypt_password = (CryptUnprotectData(encrypted_password)).decode()
decrypt_Secret = (CryptUnprotectData(encrypted_Secret)).decode()
####################generate access token####################
print(f"Start.................Script")
uri = "https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/token"
ReqTokenBody = {
    'Grant_Type'    : 'Password',
    'Username'      : 'SukhijaV@techwizard.cloud',
    'Password'      : decrypt_password,
    'client_Id'     : decrypt_clientid,
    'Client_Secret' : decrypt_Secret,
    'redirect_uri'  : RedirectUri,
    'Scope'         : 'https://graph.microsoft.com/.default',
  } 
try:
 response = requests.post(uri,data=ReqTokenBody)
 json_string = json.dumps(response.json())
 json_dict = json.loads(json_string)
 access_token =  json_dict['access_token']
except Exception as error_occured:
  print(f"Entered password: {error_occured}")
###################Now lets get the SPO list####################
head =  {'Authorization': 'Bearer ' + access_token}
response = requests.get(guri, headers=head)
json_response_string = json.dumps(response.json())
json_response_dict = json.loads(json_response_string)
listItems = json_response_dict['value']

coll_listItems=[]
response = requests.get(guri, headers=head)
json_response_string = json.dumps(response.json()) #striing values from json
json_response_dict = json.loads(json_response_string) # convert to dict
coll_listItems = json_response_dict['value'] # list
next_link = json_response_dict['@odata.nextLink']
print(f"Initial Graph request submitted")
count = 0
while True:
    try:
        count = count + 1
        print(f"Page count: {count}")
        coll_listitems_1=[]
        json_response_string = json_response_dict = None
        response = requests.get(next_link, headers=head)
        json_response_string = json.dumps(response.json())
        json_response_dict = json.loads(json_response_string)
        coll_listitems_1 = json_response_dict['value'] # list
        next_link = json_response_dict['@odata.nextLink']
        coll_listItems.extend(coll_listitems_1)
    except KeyError: 
        print(f"Key Not Found")
        coll_listitems_1=[]
        json_response_string = json.dumps(response.json())
        json_response_dict = json.loads(json_response_string)
        coll_listitems_1 = json_response_dict['value'] # list
        coll_listItems.extend(coll_listitems_1)
        break
    except Exception as error_occured:
        print(f"exception Occured{error_occured}")
