import pandas as pd
import os
import win32com.client as win32
import traceback

# import xlsxwriter

f = open("Auto Send Conf.txt") 
lines = f.readlines()

def Auto_Send():
    folder = str(lines[0].rstrip())  #第一行是文件夹路径
    # 邮箱原来在这里****email = input('Please input the email.xlsx path: ')
    # onbehalf = input('Send on behalf? Y/N: ')
    # onbehalf_name = input('Skip OR Input Email Address: ')

    try:
        if not os.path.exists(folder):
            print('No such folder!')
        else:
            email = str(lines[1].rstrip()) #第二行是邮箱list
            if not os.path.exists(email):
                print('No such file!')
            else:    
                email_excel = pd.read_excel(email,keep_default_na=False)
                filename = list(set(email_excel['File Name 文件名']))
                filename.sort()
                for name in filename:
                    pending_to = list(email_excel['Recipients 收件人'][email_excel['File Name 文件名'] == name])
                    # 有同名文件，提取多个收件人，抄送人，主题等等
                    i = len(pending_to)  # 同名文件的数量
                    for j in range(i):
                        #check File Name, no file, skip
                        if str(name) == "":
                            print('No attachment')
                            continue
                        else:
                            outlook = win32.Dispatch('Outlook.Application') #打开Outlook
                            msg = outlook.CreateItem(0)  # 0: olMailItem
                            msg.Attachments.Add(f'{str(folder)}\{str(name)}')# 附件的路径
                            #check sender, no name, skip
                            if str(list(email_excel['Sender 发件人'][email_excel['File Name 文件名'] == name])[j]) == "":
                                print(str(name) + ' must have sender 发件人')
                                continue
                            else:
                                msg.SentOnBehalfOfName = list(email_excel['Sender 发件人'][email_excel['File Name 文件名'] == name])[j]
                                #must have one receiver
                                if str(pending_to[j]) == "" and str(list(email_excel['CC 抄送人'][email_excel['File Name 文件名'] == name])[j]) == "" and str(list(email_excel['Bcc 密送人'][email_excel['File Name 文件名'] == name])[j]) == "":
                                    print(str(name) + ' must have one receiver')
                                    continue
                                else:
                                    #check receiver
                                    if str(pending_to[j]) == "":
                                        msg.to = " "
                                    else:
                                        msg.to = pending_to[j]
                                    #check CC
                                    if str(list(email_excel['CC 抄送人'][email_excel['File Name 文件名'] == name])[j]) == "":
                                        msg.cc = " "
                                    else:
                                        msg.cc = list(email_excel['CC 抄送人'][email_excel['File Name 文件名'] == name])[j]
                                    #check Bcc
                                    if str(list(email_excel['Bcc 密送人'][email_excel['File Name 文件名'] == name])[j]) == "":
                                        msg.Bcc = " "
                                    else:
                                        msg.Bcc = list(email_excel['Bcc 密送人'][email_excel['File Name 文件名'] == name])[j]
                                    #check Subject
                                    if str(list(email_excel['Subject 邮件主题'][email_excel['File Name 文件名'] == name])[j]) == "":
                                        msg.Subject = " "
                                    else:
                                        msg.Subject = list(email_excel['Subject 邮件主题'][email_excel['File Name 文件名'] == name])[j]
                                    msg.BodyFormat = 2  # 2是html格式
                                    #check Content
                                    if str(list(email_excel['Content 邮件内容'][email_excel['File Name 文件名'] == name])[j]) == "":
                                        msg.HTMLBody = " "
                                    else:
                                        msg.HTMLBody = list(email_excel['Content 邮件内容'][email_excel['File Name 文件名'] == name])[j]

                                    msg.Display()  # 显示发送邮件界面
                                    msg.Send()
                                    print(str(name) + ' ---> ' + str(pending_to[j]) + ' Sent Successfully! 发送成功')
    except Exception:
        tb = traceback.format_exc()
        print('Error Message!!!\n',tb)

#默认执行一遍
Auto_Send()

#重复执行脚本
while True:
    repeat = input("Do you want to repeat the auto send script? Y/N: ").lower()
    if repeat =="y":
        Auto_Send()
    else:
        break

os.system('pause')
#no attachment, skip
#no sender, skip
#no receiver/CC/BCC, stop
