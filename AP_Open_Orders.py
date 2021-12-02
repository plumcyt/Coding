import win32com.client
#import re
import os
#import pandas as pd
import time
import datetime
import traceback
import pytz
utc=pytz.UTC

#txt = input('Please input the file path: ')
f = open("Auto Save Conf.txt") 
lines = f.readlines()
sender = str(lines[0].rstrip()).lower() #第一行是sender
subject = str(lines[1].rstrip()).lower() #第二行是subject
mailbox = str(lines[2].rstrip()) #第三行是AI GPCGS
folder = str(lines[3].rstrip()) #第三行是AP Open Orders 
dirpath = str(lines[4].rstrip()).lower() #第四行是文件夹

latest = max([os.path.join(dirpath,d) for d in os.listdir(dirpath)], key=os.path.getctime)
newest = utc.localize(datetime.datetime.strptime(time.ctime(os.path.getctime(latest)), "%a %b %d %H:%M:%S %Y"))
days183 = newest - datetime.timedelta(days=183)

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#root_folder = outlook.Folders
#AI = root_folder["AI GPCGS"].Folders 

folder = outlook.Folders[mailbox].Folders[folder]
messages = folder.Items

	
try:
        total = len(messages)
        count1 = 0
        count2 = 0
        for email in messages:

                if email.Subject.lower() == subject and email.Sender.Name.lower() == sender and email.SentOn > newest:
                        attachments = email.Attachments
                        attachment = attachments.Item(1)
                        #attachment_name = str(attachment).lower()
                        file_date = str(email.SentOn)[0:10]
                        attachment.SaveASFile(f'{str(lines[4].rstrip())}\GPC AP Open Orders {str(email.SentOn)[0:10]}.xlsx')
                        count1 += 1
                        print(str(count1) + ' ' + 'GPC AP Open Orders ' + str(email.SentOn.date()) + ' saved')

#跳过去不处理
                #elif email.SentOn <= newest and email.SentOn > days183:
                        #count2 += 1
                        #os.system('cls')
                        #print('\r' + str(count2) + ' ' + str(email.Subject) + ' ' + str(email.SentOn)[0:10] + ' not saved',end="")

                else:
                        pass
			

except Exception:
	tb = traceback.format_exc()
	print('Error Message!!!\n',tb)
print('Finished!')
os.system('pause')
