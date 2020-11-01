"""
    .NOTES
    ===========================================================================
    Created with: 	ISE
    Created on:   	10/03/2020 1:46 PM
    Created by:   	Vikas Sukhija
    Organization:
    Filename:     	ExportADAttributes.py
    ===========================================================================
    .DESCRIPTION
    This script can run and export AD attributes for all Enabled users in AD
    Libraries used: ldap3
"""
import csv
from ldap3 import Server, \
    Connection, \
    AUTO_BIND_NO_TLS, \
    SUBTREE


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
log_file = create_file("logs", "export_ad", "log")
report_file = create_file("report", "export_ad", "csv")
base_dn = "DC=labtest,DC=com"
domain_controller = "DC"
admin_userid = "labtest\\vikas"
admin_password = "Password"
obj_filter = "(&(objectClass=user)(objectCategory=person)(!(useraccountcontrol:1.2.840.113556.1.4.803:=2)))"

write_log("start.....script", log_file)
write_log("Querying..........AD", log_file)

with Connection(Server(domain_controller, port=636, use_ssl=True),
                auto_bind=AUTO_BIND_NO_TLS,
                read_only=True,
                check_names=True,
                user=admin_userid, password=admin_password) as c:
    c.extend.standard.paged_search(search_base=base_dn,
                                   search_filter=obj_filter,
                                   search_scope=SUBTREE,
                                   attributes={'sAMAccountName', 'userPrincipalName', 'mail', 'mobile', 'title',
                                               'department'},
                                   get_operational_attributes=True,
                                   paged_size=1500,
                                   generator=False)

write_log("Fetched Ad Attributes", log_file)
write_log("Export to CSV file", log_file)

with open(report_file, mode='w', newline='', encoding='utf-8') as csv_file:
    fieldnames = ['sAMAccountName', 'userPrincipalName', 'mail', 'mobile', 'title',
                  'department']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()
    for user in c.entries:
        write_log("exporting..." + str(user["sAMAccountName"]), log_file)
        writer.writerow(
            {'sAMAccountName': user["sAMAccountName"], 'userPrincipalName': user["userPrincipalName"],
             'mail': user["mail"], 'mobile': user["mobile"], 'title': user["title"],
             'department': user["department"]})

write_log("Script.......Finished", log_file)
