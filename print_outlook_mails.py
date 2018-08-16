# -*- coding: utf-8 -*-
"""
Created on Wed Aug  8 09:54:17 2018

@author: ejiandi
"""

import datetime
from win32com.client.gencache import EnsureDispatch as Dispatch
outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
Accounts = mapi.Folders  # 根级目录（邮箱名称，包括Outlook读取的存档名称）
for Account_Name in Accounts:
    #print(' >> 正在查询的帐户名称：',Account_Name.Name,'\n')
    Level_1_Names = Account_Name.Folders  # 一级目录集合（与inbox同级）
    for Level_1_Name in Level_1_Names:
        # 首先，向MySQL提交一级目录的邮件
        #print(' - 正在查询一级目录：' , Level_1_Name.Name)
        # 然后，判断当前查询的一级邮件目录是否有二级目录（若有多级目录，可以参考此段代码)
        if Level_1_Name.Folders: 
            Level_2_Names = Level_1_Name.Folders  # 二级目录的集合（比如，自建目录的子集）
            for Level_2_Name in Level_2_Names:
                #print(' - - 正在查询二级目录：' , Level_1_Name.Name , '//' , Level_2_Name.Name)
                if (Level_2_Name.Name == 'offshore'):
                    Mail_2_Messages = Level_2_Name.Items  # 二级目录的邮件集合
                else:
                    Mail_2_Messages = ''
                for yy in Mail_2_Messages:  # xx = 'mail'  # 开始查看单个邮件的信息
                    Root_Directory_Name_2 = Account_Name.Name # 记录根目录名称
                    Level_1_FolderName_2 = Level_1_Name.Name # 记录一级目录名称
                    Level_2_FolderName_2 = Level_2_Name.Name # 记录二级目录名称
                    if (hasattr(yy, 'ReceivedTime')):
                        R = yy.ReceivedTime
                        R = datetime.date(R.year, R.month, R.day)   #生成收件日期(datetime.date的格式)
                        datedelta = datetime.date.today() - R       #距离今天有多久?(datetime.timedelta的格式)
                        if (datedelta < datetime.timedelta(60)) and (datedelta > datetime.timedelta(0)):
                            weeks = datetime.date.isocalendar(yy.ReceivedTime)  #转换成ISO日历格式
                            weeks = weeks[1]                                    #ISO日历格式的第二项就是第几周的信息
                        else:
                            weeks = ''
                    else:
                        R = ''
#                    if (hasattr(yy, 'SenderName')):  # 发件人
#                        SenderName_2 = yy.SenderName
#                    else:
#                        SenderName_2 = ''
#                    if (hasattr(yy, 'Subject')):  # 主题
#                        Subject_2 = yy.Subject
#                    else:
#                        Subject_2 = ''
                    if (hasattr(yy, 'ConversationTopic')):  # 会话主题
                        ConversationTopic_2 = yy.ConversationTopic
                        ConversationTopic_2 = ConversationTopic_2.split(']',1)  #这个分割用的字符']'消失了
                        if (yy.ConversationTopic[5] == '#'):
                            RTT = ConversationTopic_2[0]
                        else:
                            RTT = ''
                    else:
                        ConversationTopic_2 = ''
                    if (hasattr(yy, 'Body')):  # 邮件正文内容
                        EmailBody_2 = yy.Body
                        EmailBody_2 = EmailBody_2.splitlines(True)  #这个True会保留换行符,后面截取到最后一个字符的时候没有它就把内容最后一个字符给切没了
                        subject = EmailBody_2[0]
                        subject = subject[14:-1]
                        if (RTT == ''):
                            RTT = EmailBody_2[4]
                            RTT = '[' + RTT[0:4] + ' ' + RTT[5:12]
                            subject = EmailBody_2[6]
                            subject = subject[7:-1]
                    else:
                        EmailBody_2 = ''
                    if (weeks != '') :
                        print ('\tW' + str(weeks), '\t', RTT + ']' + subject,'\t')
        else:
            pass

# 结尾

print ('\n',' >> Done!')
