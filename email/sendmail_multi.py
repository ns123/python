#! /usr/bin/env python  
# coding:utf-8

from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr

import smtplib
#import datetime
import os
import codecs
import ConfigParser

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, 'utf-8').encode(), \
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))

class Db_Connector:
    def __init__(self, config_file_path):
        self.cf = ConfigParser.ConfigParser()
        self.cf.readfp(codecs.open(config_file_path,"r","utf-8-sig"))

    def GetEmailList(self):
        s=self.cf.sections()
        return s

    def GetCustomerOption(self,option):
        o=self.cf.options(option)
        return o

    def GetCustomerItems(self,option):
        out=self.cf.items(option)
        return out

    def get(self, field, key):
        result = ""
        try:
            result = self.cf.get(field, key)
        except:
            result = ""
        return result

def JudgeDirExist(DirIn):#判断传入的列表里，文件路径是不是存在，不存在返回错误路径
    if os.path.exists(DirIn) == False:
        input("Please check out the ToEmailList.ini %s !"% DirIn)
    else:
        return True
'''
def FindTheLastFile(base_dir):#根据传入的目录名找到创建时间为最后的文件，并返回其完整路径
    DirInList = os.listdir(base_dir)
    filelist = []#文件列表
    timelist = []#时间列表
    for i in range(0, len(DirInList)):
        path = os.path.join(base_dir,DirInList[i])#path=目录+文件名
        if os.path.isfile(path):
            filelist.append(DirInList[i])

    for i in range(0, len(filelist)):
        path = os.path.join(base_dir, filelist[i])
        if os.path.isdir(path):
            continue
        #timestamp = os.path.getmtime(path)
        #print timestamp
        timelist.append(os.path.getctime(path))#getctime创建时间

    TheLastFileName = filelist[timelist.index(max(timelist))]

    return os.path.join(base_dir,TheLastFileName)#path=目录+文件名
'''
class FindTheLastFile(object):
    def __init__(self,base_dir):
        self.bdir=base_dir
        self.file_name=''
        pass

    def find_file_name(self):
        self.TheLastFileName = ''
        self.DirInList = os.listdir(self.bdir)
        filelist = []
        timelist = []
        for i in range(0, len(self.DirInList)):
            path = os.path.join(self.bdir,self.DirInList[i])#path=目录+文件名
            if os.path.isfile(path):
                filelist.append(self.DirInList[i])

        for i in range(0, len(filelist)):
            path = os.path.join(self.bdir, filelist[i])
            if os.path.isdir(path):
                pass

            else:
                timelist.append(os.path.getctime(path))#getctime创建时间
        self.TheLastFileName = filelist[timelist.index(max(timelist))]
        self.file_name = self.TheLastFileName
        return self.TheLastFileName

    def Whole_Path(self):
        return os.path.join(self.bdir,self.file_name)#path=目录+文件名



if __name__ == '__main__':

    from_addr = raw_input('From: ')
    password = raw_input('Password: ')
    smtp_server = raw_input('SMTP server(like: smtp.qq.com): ')

    d = Db_Connector("./ToEmailList.ini")
    to_addr_list=d.GetEmailList()

    times=0
    print 'start'

    for to_addr in to_addr_list:
        to_title = d.get(to_addr,'Title')
        to_main = d.get(to_addr,'Text')
        to_dir = d.get(to_addr,'Dir')

        
        JudgeDirExist(to_dir)#判断一下是不是有不存在的路径
        dirs = FindTheLastFile(to_dir)
        dirs.find_file_name()


        msg = MIMEMultipart()
        msg['From'] = _format_addr(u'发送自 <%s>' % from_addr)
        msg['To'] = _format_addr(u'管理员 <%s>' % to_addr)
        msg['Subject'] = Header(u'%s' %to_title, 'utf-8').encode()
        # add MIMEText:
        msg.attach(MIMEText(u'%s' % to_main, 'plain', 'utf-8'))

        # add file:
        with open(dirs.Whole_Path(), 'rb') as f:
            mime = MIMEBase('Excel', 'xlsx', filename=dirs.file_name)
            mime.add_header('Content-Disposition', 'attachment', filename=dirs.file_name)
            mime.add_header('Content-ID', '<0>')
            mime.add_header('X-Attachment-Id', '0')
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            msg.attach(mime)

        server = smtplib.SMTP(smtp_server, 25)
        #server.set_debuglevel(1)
        server.login(from_addr, password)
        server.sendmail(from_addr, [to_addr], msg.as_string())
        server.quit()

        times+=1
        print times,' Send Success'
        if times==len(to_addr_list):
            print  'All successful!'

    try:
        a=input("Please <enter>")
    except Exception:
        print 'bye'
