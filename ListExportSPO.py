"""
    .NOTES
    ===========================================================================
    Created with: 	PyCharm
    Created on:   	10/06/2020 1:46 PM
    Created by:   	Vikas Sukhija
    Organization:
    Filename:     	ListExportSPO.py
    ===========================================================================
    .DESCRIPTION
    This script can run and export spo list items
    Libraries used: Office365-REST-Python-Client
"""
import csv
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext


# Load generic functions
# Write_log
def write_log(message, file_path, severity="Information"):
    sev_types = ['Information', 'Warning', 'Error']
    if severity not in sev_types:
        raise ValueError("Invalid severity type. Expected one of: %s" % sev_types)

    import datetime
    d = datetime.datetime.today()
    d_formatdate = d.strftime('%m-%d-%Y-%I-%M-%S')
    t_green = '\033[32m'  # Green Text
    t_red = '\033[31m'  # Red Text
    t_yellow = '\033[33m'  # Red Text

    import os
    if severity == "Information":
        if not os.path.isfile(file_path):
            raise ValueError("Invalid operation: %s not found" % file_path)
        else:
            f = open(file_path, "a")
            f.write("\n")
            f.write(d_formatdate + " | " + message + " | " + severity + "|")
            print(t_green + d_formatdate + " | " + message + " | " + severity + "|" + t_green)
            f.close()

    if severity == "Warning":
        if not os.path.isfile(file_path):
            raise ValueError("Invalid operation: %s not found" % file_path)
        else:
            f = open(file_path, "a")
            f.write("\n")
            f.write(d_formatdate + " | " + message + " | " + severity + "|")
            print(t_yellow + d_formatdate + " | " + message + " | " + severity + "|" + t_yellow)
            f.close()

    if severity == "Error":
        if not os.path.isfile(file_path):
            raise ValueError("Invalid operation: %s not found" % file_path)
        else:
            f = open(file_path, "a")
            f.write("\n")
            f.write(d_formatdate + " | " + message + " | " + severity + "|")
            print(t_red + d_formatdate + " | " + message + " | " + severity + "|" + t_red)
            f.close()


# Write_log


# create_file
def create_file(folder_name, name, ext):
    import datetime
    d = datetime.datetime.today()
    d_formatdate = d.strftime('%m-%d-%Y-%I-%M-%S')
    filename = folder_name + "\\" + name + "_" + d_formatdate + "_." + ext
    import os
    if not os.path.exists(folder_name):
        os.mkdir(folder_name)
    f = open(filename, "x")
    f.close()
    return filename


# create_file

# log and variables
log_file = create_file("logs", "spo_list_item", "log")
report_file = create_file("report", "spo_list_item", "csv")
site_url = "https://techwizard.sharepoint.com/sites/TeamSIte/"
sp_list = "DL Modification"
admin_user = "jtechwizard@techwizard.cloud"
admin_password = "password"

write_log("start.....script", log_file)

try:
    ctx = ClientContext(site_url).with_credentials(UserCredential(admin_user, admin_password))
    sp_lists = ctx.web.lists
    s_list = sp_lists.get_by_title(sp_list)
    l_items = s_list.get_items()
    ctx.load(l_items)
    ctx.execute_query()
    write_log("Fetched list items", log_file)
    with open(report_file, mode='w', newline='', encoding='utf-8') as csv_file:
        fieldnames = ['ID', 'Title', 'Check']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for item in l_items:
            write_log("exporting..." + str(item.properties['ID']), log_file)
            writer.writerow({'ID': item.properties['ID'], 'Title': item.properties['Title'],
                             'Check': item.properties['Check']})
except:
    write_log("Exception ....... caught", log_file, severity="Error")

write_log("Script.......Finished", log_file)
