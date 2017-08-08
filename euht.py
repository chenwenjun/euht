# -*- coding: utf-8-*-

import time,threading
import os
import telnetlib
import xlrd
from xlutils.copy import copy;


print u'开始统计'


 # 配置选项

router_status = False
router_username = b'admin'
router_password = b'!iK9u5!!fUkjKou!9klIUf!23mjJD123r!87k!'  
router_term = b'>'

#scan_flag = 1  代表全局扫描，scan_flag = 0 有开通扫描，要根据excel来决定
scan_flag = 1

cap_status = False
cap_term = b'~$'
cap_username = b'root'
cap_num = 2
cap_soft_version = 'b36_u'
cap_upgrade_version = 'b37'


xls_path = u'E:/back/work/EUHT基站.xlsx'
xls_new = u'E:/back/work/EUHT基站_new.xlsx'
xls_routerip_col = 6
xls_capip_col = 7
xls_village_row = 1
xls_capnum_col = 11



xls_cap1_channel = 12

xls_router_status = 15
xls_router_status 
xls_cap1_status = 16
xls_cap1_version = 25

xls_cap1_txpower = 31

class CityInfo (object):

    def __init__ (self,name,village_row=1):
        self.name = name
        self.village_row = village_row



class Router (object):

    def __init__ (self,router_ip,cap_num,capip_list):
        self.router_ip = router_ip
        self.cap_num = cap_num
        self.capip_list = capip_list
        self.tn = 0

    def get_router_ip (self):
        return self.router_ip


    def get_telnet_router (self):
        return self.tn

    def telnet_router (self,port,retry,timeout):
        try_count = 0
        while  (try_count < retry):
            try:
                tn = telnetlib.Telnet(self.router_ip,port,timeout)
                tn.set_debuglevel(3) 
                telnet_rt_log = tn.read_until(b"Username: ",5)
                tn.write(router_username + b"\n")
                telnet_rt_log = tn.read_until(b"Password: ",5)
                tn.write(router_password + b"\n")
                if (telnet_rt_log.find(b"Password: ") > -1):
                    print "telnet router: %s successfully " % self.router_ip
                    self.tn  = tn
                    return True
                else:
                    try_count = try_count + 1
                    
            except Exception as e:
                try_count = try_count + 1
                print (e)
                print ("telnet router fail , errcode = %d "  % try_count) 
                
            if (try_count == retry):
                return False

class CAP (object):

    def __init__ (self,ip,tn):
        self.ip = ip
        self.tn = tn


    def telnet_cap (self,timeout,retry):
        try_count = 0
        while (try_count < retry):
            try:    
                tn = self.tn
                tn.write(b"telnet " + self.ip + b"\n")
                cap_login_info = tn.read_until(b"login: ",timeout)
                if (cap_login_info.decode().find('login') > 0):
                    print "login cap ok "
                    tn.write(b"root\n")
                    return True
                else:
                    tn.write(chr(3).encode('utf-8')+b"\n")
                    print ("login cap fail %d " % try_count)
                    try_count = try_count + 1
                    
            except Exception as e:
                try_count = try_count + 1
                print e
                print "telnet cap " + cap_ip.decode() + " exception"
            
            if (try_count == retry):
                return False


    def get_cap_version (self):
        try:
            tn = self.tn
            tn.read_until(cap_term,10)
            tn.write(b"cat /etc/buildid.conf \n")
            info = tn.read_until(b".bit",3) 
            capverinfo = info.decode()
            capver = capverinfo[capverinfo.find('basic.b')+6:capverinfo.find('basic.b')+9]

            if (capver.find(cap_upgrade_version) == -1):
                print ("low vesion")
            return capver
        except Exception as e:
            print (e)
            print ("get cap soft vesrion fail ")
        return "unkown"

    def get_cap_channel (self):
        try:
            tn = self.tn
            tn.read_until(cap_term,10)
            tn.write(b"cat /misc/iniCfg | grep channel  \n")
            time.sleep(3)
            info = tn.read_very_eager() 
            capchaninfo = info.decode()
            print "--------111"+capchaninfo+"222----"
            capver = capchaninfo[capchaninfo.find('=')+1:len(capchaninfo)-16]

            #if (capver.find(cap_upgrade_version) == -1):
                #print ("low vesion")
            return capver
        except Exception as e:
            print e
            print "get cap channel vesrion fail "
        return "unkown"

    def get_cap_txpower (self):
        try:
            tn = self.tn
            tn.read_until(cap_term,10)
            tn.write(b"cat /misc/iniCfg | grep txPower  \n")
            time.sleep(3)
            info = tn.read_very_eager() 
            capchaninfo = info.decode()
            print ("--------111"+capchaninfo+"222----")
            capver = capchaninfo[capchaninfo.find('=')+1:len(capchaninfo)-16]

            #if (capver.find(cap_upgrade_version) == -1):
                #print ("low vesion")
            return capver
        except Exception as e:
            print (e)
            print ("get cap channel vesrion fail ")
        return "unkown"

    def exit_cap (self):
        tn  = self.tn
        tn.write(b"exit\n")
        tn.read_until(router_term,10)



def work (city,sheet,newsheet):
    print '-------------thread %s is running...\n' % threading.current_thread().name.decode("utf-8")
    xls_village_row = city.village_row


    newsheet.write(0, xls_router_status, u"路由器状态")

    newsheet.write(0, xls_router_status+1, u"CAP1状态")
    newsheet.write(0, xls_router_status+2, u"CAP2状态")
    newsheet.write(0, xls_router_status+3, u"CAP3状态")
    '''
    newsheet.write(0, xls_cap1_version, "CAP1版本")
    newsheet.write(0, xls_cap1_version+1, "CAP2版本")
    newsheet.write(0, xls_cap1_version+2, "CAP3版本")
    '''
    newsheet.write(0, xls_cap1_channel, u"CAP1频点")
    newsheet.write(0, xls_cap1_channel+1, u"CAP2频点")
    newsheet.write(0, xls_cap1_channel+2, u"CAP3频点")

    """
    newsheet.write(0, xls_cap1_txpower, "CAP1 power")
    newsheet.write(0, xls_cap1_txpower+1, "CAP2power")
    newsheet.write(0, xls_cap1_txpower+2, "CAP3power")
    """

    while (xls_village_row < sheet.nrows):
        print sheet.cell_value(xls_village_row, 1),sheet.cell_value(xls_village_row, 2),sheet.cell_value(xls_village_row, 3),sheet.cell_value(xls_village_row, 4)
        routerip = sheet.cell_value(xls_village_row, xls_routerip_col)
        print 'xls_village_row ', xls_village_row
        print "routerip " , routerip
        cap_num = sheet.cell_value(xls_village_row, xls_capnum_col)
        
        if (not routerip):
            xls_village_row = xls_village_row + 1       
            continue
        else:
            if (not cap_num):
                if (scan_flag == 0):
                    xls_village_row = xls_village_row + 1  
                    print "no cap !"                   
                    continue  
                else:
                    cap_num = 3 #每个站点CAP数最大为3  

        cap_num = int (cap_num)
        print "cap num %d " % cap_num
        if (cap_num > 0):   
            capip_list = [ sheet.cell_value(xls_village_row, xls_capip_col).encode('utf-8'), sheet.cell_value(xls_village_row, xls_capip_col+1).encode('utf-8'),sheet.cell_value(xls_village_row, xls_capip_col+2).encode('utf-8')]
            print capip_list
        else:
            continue
        
        router = Router (routerip,cap_num,capip_list)

        router_status = router.telnet_router (2306,3,3)       
            
        if (router_status):
            newsheet.write(xls_village_row, xls_router_status, "on")
            telnet_cap_num = 0
            while (telnet_cap_num < cap_num):
                cap = CAP (router.capip_list[telnet_cap_num],router.tn)
                cap_status = cap.telnet_cap (5,3)
                if (cap_status == True):
                    """
                    newsheet.write(xls_village_row, xls_cap1_status + telnet_cap_num, "online")
                    cap_soft_version =  get_cap_version()
                    print ("************* CAP%d version "% (telnet_cap_num+1) + cap_soft_version + " *************" )
                    newsheet.write(xls_village_row, xls_cap1_version + telnet_cap_num, cap_soft_version)
                     """
                    cap_channel_info = cap.get_cap_channel ()
                    print "---------cap "+cap_channel_info
                    newsheet.write(xls_village_row, xls_cap1_channel + telnet_cap_num, cap_channel_info)
                   
                    """
                    cap_txpower_info = get_cap_txpower ()
                    print ("---------cap "+cap_txpower_info)
                    newsheet.write(xls_village_row, xls_cap1_txpower + telnet_cap_num, cap_txpower_info)   
                    """
                    newsheet.write(xls_village_row, xls_cap1_status + telnet_cap_num, "on")            
                    cap.exit_cap ()
                else:
                    newsheet.write(xls_village_row, xls_cap1_status + telnet_cap_num, "off")
                telnet_cap_num = telnet_cap_num + 1
                print "telnet_cap_num %d "  % telnet_cap_num
            #退出工业路由器
            #tn.write(b"exit\n")
        else:
            newsheet.write(xls_village_row, xls_router_status, "off")
        xls_village_row = xls_village_row + 1


#读取数据
book = xlrd.open_workbook(xls_path)
print "\nThe number of worksheets in %s is %d" %(xls_path, book.nsheets)


for sheet_names in book.sheet_names():
    print sheet_names,

print "now working ........ "  


newbook = copy(book)

threads = []
for i in range (1,14):
    #print sheet.name, sheet.nrows, sheet.ncols
    sheet = book.sheet_by_index(i)
    newsheet = newbook.get_sheet(i)
    #print sheet.name,newsheet.name
    #print newsheet.name, newsheet.nrows, newsheet.ncols
    city =  CityInfo(sheet.name, xls_village_row)
    t = threading.Thread(target=work, name=city.name.encode("utf-8"),args=(city,sheet,newsheet))
    threads.append (t)
    t.start()

for t in threads:
    t.join ()

try:
    os.remove(xls_new)
except Exception as e:
    pass
newbook.save(xls_new)
print "save  ok"
