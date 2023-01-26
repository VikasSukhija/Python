"""
    .NOTES
    ===========================================================================
    Created with: 	vscode
    Created on:   	12/21/2022 1:46 PM
    Created by:   	Vikas Sukhija
    Organization:
    Filename:     	RemoveSpecificADGroupFromUsers.py
    ===========================================================================
    .DESCRIPTION
    Check all users under research OUs, for users that are members of groups starting with VS-PWGP- , remove the groups.
"""
import configparser
import os
from ldap3 import Server, Connection, AUTO_BIND_NO_TLS, SUBTREE, ALL_ATTRIBUTES #for active directory
from ldap3.extend.microsoft.addMembersToGroups import ad_add_members_to_groups as addUsersInGroups
from ldap3.extend.microsoft.removeMembersFromGroups import ad_remove_members_from_groups as removeUsersInGroups
import xml.etree.ElementTree as ET
from DPAPI  import CryptUnprotectData
from vsadmin import create_file,write_log, send_mail, recycle_logs
########################logs and Variables#######################
log_file = create_file("logs", "RemoveSpecificADGroupFromUsers", "log")
configfile = "E:\\scheduledscripts\\identity\\Config.ini"
countofchanges = 20
sender_address = "DoNotReply@labtest.com"

groupsuffix = "VS-PWGP-*"
rshOUstring = "Research"

config = configparser.ConfigParser()
config.read(configfile)
err_rcpt_to = config['Admin']['ErrorEmail']
smtp_server = config['Admin']['SMTPServer']
base_dn = config['Admin']['VSdomaindn']
dc_1 = os.environ['logonserver']
dc_1 = dc_1.replace("\\","")

current_directory = os.getcwd()
logs_folder = current_directory + "\\" + "logs"

obj_filter = "(&(objectClass=user)(objectCategory=person)(!(useraccountcontrol:1.2.840.113556.1.4.803:=2))(|(proxyAddresses=*)(employeeID=*)))"
group_filter = "(&(objectClass=group)(Name=VS-PWGP-*))"
#########################fetch admin credentials#################
try:
    admin_user = config['ServiceAccount']['UserID']
    Password1 = config['ServiceAccount']['Password']
    Password = CryptUnprotectData(bytes.fromhex(Password1)).decode().replace('\x00', '')
except Exception as error_occured:
    write_log("exception Occured" + str(error_occured), log_file,"Error")
    send_mail(smtp_server, sender_address,err_rcpt_to,"Error - RemoveSpecificADGroupFromUsers",str(error_occured))
    quit()
#####################Start the main code###################################
try:
    write_log(f"Start............Script", log_file)
    with Connection(Server(dc_1, port=636, use_ssl=True),auto_bind=AUTO_BIND_NO_TLS,read_only=True,check_names=True,user=admin_user, password=Password) as c:
                        c.extend.standard.paged_search(search_base=base_dn,search_filter=obj_filter,search_scope=SUBTREE,attributes={'sAMAccountName', 'distinguishedName', 'memberof'},get_operational_attributes=True,paged_size=1500,generator=False)
    All_Users = c.entries
    write_log("Loaded lad query " + str(len(All_Users)), log_file)

    f_All_Users = filter(lambda User: "TEST-" not in str(User.sAMAccountName), All_Users)
    f_Users = list(f_All_Users)
    f_mfg_users = filter(lambda User: rshOUstring in str(User.distinguishedName), f_Users)
    mfg_users = list(f_mfg_users)
    write_log("Finished loading AD Users " + str(len(mfg_users)), log_file)

    with Connection(Server(dc_1, port=636, use_ssl=True),auto_bind=AUTO_BIND_NO_TLS,read_only=True,check_names=True,user=admin_user, password=Password) as c:
                        c.extend.standard.paged_search(search_base=base_dn,search_filter=group_filter,search_scope=SUBTREE,attributes={'sAMAccountName', 'distinguishedName'},get_operational_attributes=True,paged_size=1500,generator=False)
    All_groups = c.entries
    write_log("Finished loading Groups " + str(len(All_groups)), log_file)
except Exception as error_occured:
    write_log("exception Occured" + str(error_occured), log_file,"Error")
    send_mail(smtp_server, sender_address,err_rcpt_to,"Error - RemoveSpecificADGroupFromUsers",str(error_occured))
    quit()
###########################fetch users with PWGP group############################
try:
    write_log("Start fetching users with PWGP groups", log_file)
    groupcoll = []
    for user in mfg_users:
        for group in All_groups:
            ADGroupDN = group['distinguishedName']
            if ADGroupDN in user['memberof']:
                 groupcoll.append(user)
                 write_log("Found User " + str(user['sAMAccountName']) , log_file)
                 break
    write_log("Found Users " + str(len(groupcoll)) , log_file)
except Exception as error_occured:
    write_log("exception Occured" + str(error_occured), log_file,"Error")
    send_mail(smtp_server, sender_address,err_rcpt_to,"Error - RemoveSpecificADGroupFromUsers",str(error_occured))
    quit()
##################################start removing the PWGP groups###################
try:
    if len(groupcoll) > 0 and len(groupcoll) < countofchanges:
        for u in groupcoll:
            user_dn = u['distinguishedName']
            usersam = u['sAMAccountName']
            write_log("Processing......." + str(usersam), log_file)
            for group in All_groups:
                ADGroupDN = group['distinguishedName']
                if ADGroupDN in u['memberof']:
                    write_log(f"Found group {usersam } - {user_dn} - {ADGroupDN}" , log_file)
                    with Connection(Server(dc_1, port=636, use_ssl=True),auto_bind=AUTO_BIND_NO_TLS,read_only=False,check_names=True,user=admin_user, password=Password) as c:
                        removeUsersInGroups(c, str(user_dn), str(ADGroupDN),fix=True)
    elif len(groupcoll) > countofchanges:
        write_log("Number of requests are " + str(len(groupcoll)) + "more than " + str(countofchanges), log_file)
        send_mail(smtp_server, sender_address,err_rcpt_to,"Error - Number of requests are More - RemoveSpecificADGroupFromUsers","Number of requests are " + str(len(groupcoll)) + "more than " + str(countofchanges))
except Exception as error_occured:
    write_log("exception Occured" + str(error_occured), log_file,"Error")
    send_mail(smtp_server, sender_address,err_rcpt_to,"Error - RemoveSpecificADGroupFromUsers",str(error_occured))
##############################Recycle logs###########################################
write_log("Script.........Finished", log_file)
recycle_logs(logs_folder,60,True)
logfile_path = os.path.join(current_directory,log_file)
send_mail(smtp_server, sender_address,err_rcpt_to,"Log - RemoveSpecificADGroupFromUsers","Log - RemoveSpecificADGroupFromUsers",cc_list=None, bcc_list=None, attachment_path=logfile_path)

####################Script Finish#####################################################



