import configparser
import pysftp
from DPAPI  import CryptUnprotectData
# Loads .ssh/known_hosts 
configfile = "D:\\scripts\\Config.ini"
sender_address = "DoNotReplyPy@labtest.com"
sftp_server = "sftp.Vendor.com"
private_key = "D:\\scripts\\Privatekey_1"
try:
    config = configparser.ConfigParser()
    config.read(configfile)
    err_rcpt_to = config['master']['ErrorEmail']
    smtp_server = config['master']['SMTPServer']
    v_user = config['Vendor']['UserID']
    vPassword1 = config['Vendor']['Password']
    vPassword = CryptUnprotectData(bytes.fromhex(vPassword1)).decode().replace('\x00', '')
except Exception as error_occured:
    quit()
cnopts = pysftp.CnOpts()
hostkeys = None
if cnopts.hostkeys.lookup(sftp_server) == None:
    print("New host - will accept any host key")
    # Backup loaded .ssh/known_hosts file
    hostkeys = cnopts.hostkeys
    # And do not verify host key of the new host
    cnopts.hostkeys = None

with pysftp.Connection(host=sftp_server, username=v_user, private_key=private_key,  private_key_pass=vPassword, cnopts=cnopts) as sftp:
    if hostkeys != None:
        print("Connected to new host, caching its hostkey")
        hostkeys.add(
            sftp_server, sftp.remote_server_key.get_name(), sftp.remote_server_key)
        hostkeys.save('E:\\Scripts\\Publickey_1.pub')