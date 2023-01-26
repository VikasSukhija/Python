"""
    .NOTES
    ===========================================================================
    Created with: 	vscode
    Created on:   	2/3/2022 1:46 PM
    Created by:   	Vikas Sukhija
    Organization:
    Filename:     	vsadmin.py
    ===========================================================================
    .DESCRIPTION
    Module for python functions
"""
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
# send_mail
def send_mail(smtp_server, from_address, to_list, subject, body, cc_list=None, bcc_list=None, attachment_path=None):
    import smtplib
    import os
    from email.message import EmailMessage
    message = EmailMessage()
    message["From"] = from_address
    message["To"] = to_list
    message["Subject"] = subject
    message.set_content(body)
    if cc_list is not None:
        message["CC"] = cc_list
    if bcc_list is not None:
        message["BCC"] = bcc_list
    if attachment_path is not None:
        filename = os.path.basename(attachment_path)
        open_file = open(attachment_path, "rb")
        message.add_attachment(open_file.read(), maintype='multipart', subtype='mixed; name=%s' % filename, filename=filename)
    s = smtplib.SMTP(smtp_server)
    s.send_message(message)
    s.quit()
# send_mail

#recycle_logs
def recycle_logs(folder_path,limit,delete=False):
    import os
    import time
    from datetime import datetime as dt,timedelta
    current_datetime = dt.now()
    days_limit = timedelta(days=limit)
    file_list = os.listdir(folder_path)
    for file in file_list:
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            last_time = os.path.getmtime(file_path)
            last_modificationTime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(last_time))
            prior_time = (current_datetime  - days_limit).strftime('%Y-%m-%d %H:%M:%S')
            if last_modificationTime < prior_time:
                print("Delete file " + file_path + " " + last_modificationTime)
                if(delete == True):
                    os.remove(os.path.join(folder_path,file))
#recycle_logs

