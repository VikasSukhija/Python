import configparser
import pysftp
from DPAPI  import CryptUnprotectData
# Loads .ssh/known_hosts 
configfile = "D:\\scripts\\Config.ini"
sender_address = "DoNotReplyPy@labtest.com"
sftp_server = "sftp.vender.com"
private_key = "D:\\scripts\\Privatekey_1"
public_key = "E:\\Scripts\\Publickey_1.pub"
try:
    config = configparser.ConfigParser()
    config.read(configfile)
    err_rcpt_to = config['master']['ErrorEmail']
    smtp_server = config['master']['SMTPServer']
    v_user = config['Vender']['UserID']
    vPassword1 = config['Vender']['Password']
    vPassword = CryptUnprotectData(bytes.fromhex(vPassword1)).decode().replace('\x00', '')
except Exception as error_occured:
    quit()
################upload to sftp#####################
try:
  cnopts = pysftp.CnOpts(knownhosts=public_key)
  conn = pysftp.Connection(host=sftp_server, username=v_user, private_key=private_key,  private_key_pass=vPassword, cnopts=cnopts)
  conn.put('D:\\temp\\UXBadgeAutomation.csv','root/import/vender.csv')
  conn.close()
except Exception as error_occured:
    quit()