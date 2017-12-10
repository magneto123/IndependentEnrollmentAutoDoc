#!/usr/bin/python
# encoding: utf-8
# -*- coding: utf8 -*-
"""
Created by PyCharm.
File:               LinuxBashShellScriptForOps:getNetworkStatus.py
User:               Guodong
Create Date:        2016/11/2
Create Time:        16:20
 
show Windows or Linux network Nic status, such as MAC address, Gateway, IP address, etc
 
# python getNetworkStatus.py
Routing Gateway:               10.6.28.254
Routing NIC Name:              eth0
Routing NIC MAC Address:       06:7f:12:00:00:15
Routing IP Address:            10.6.28.28
Routing IP Netmask:            255.255.255.0
 """
import os
import sys
import wmi#获取CPU、主板、硬盘等信息
import base64#加密算法
import platform
 
try:
    import netifaces
except ImportError:
    try:
        command_to_execute = "pip install netifaces || easy_install netifaces"
        os.system(command_to_execute)
    except OSError:
        print "Can NOT install netifaces, Aborted!"
        sys.exit(1)
    import netifaces

def GetHardwareInfo():
    '''
    function : Get Mac Address
    return the upper format string
    '''
    #获取系统类型
    ostype = platform.platform()
    
    #routingGateway = netifaces.gateways()['default'][netifaces.AF_INET][0]
    routingNicName = netifaces.gateways()['default'][netifaces.AF_INET][1]
    for interface in netifaces.interfaces():
        if not interface == routingNicName:#在vista之后的系统中要去掉not
            routingNicMacAddr = netifaces.ifaddresses(interface)[netifaces.AF_LINK][0]['addr']
            if len(routingNicMacAddr) == 17:
                break
            try:
                routingIPAddr = netifaces.ifaddresses(interface)[netifaces.AF_INET][0]['addr']
                # TODO(Guodong Ding) Note: On Windows, netmask maybe give a wrong result in 'netifaces' module.
                routingIPNetmask = netifaces.ifaddresses(interface)[netifaces.AF_INET][0]['netmask']
            except KeyError:
                pass
    #cpu 序列号
    c = wmi.WMI()
    cpu = c.Win32_Processor()[0]
    cpuid = cpu.ProcessorId.strip()
    #硬盘序列号
    if 'Windows-XP' in ostype:
        physical_disk = c.Win32_PhysicalMedia()[0]
    else:
        physical_disk = c.Win32_DiskDrive()[0]#windows vista之后的系统
    diskid = physical_disk.SerialNumber.strip()

    #主板序列号
    board_id = c.Win32_BaseBoard()[0]
    boardid = board_id.SerialNumber.strip()

    return cpuid.upper(),routingNicMacAddr.upper(),boardid.upper(),diskid.upper()


def WriteHardwareInfotoFile():
    '''
    功能：将硬件信息写入文本文件
    
    '''
    cpuid,macaddr,boardid,diskid = GetHardwareInfo()
    allinfo = cpuid + '\n' + macaddr + '\n' + boardid + '\n' + diskid
    #写入文件
    f1 = open(cpuid+'.txt', 'w')#
    f1.writelines(allinfo)
    f1.close()
    
def StrEncrypt(infostr):
    '''
    功能：对输入的字符串加密，加密算法为先将其中的数字全部加1然后字符串倒序，最后采用base64方法加密
    '''
    #先对字符串中所有的数字加1
    info = ''
    for i in range(len(infostr)):
        if infostr[i] >= '0' and infostr[i] < '9':
            info += str(int(infostr[i]) + 1)
        elif infostr[i] == '9':
            info += '0'
        else:
            info += infostr[i]
    #对字符串倒序
    info = info[::-1]
    
    return base64.encodestring(info)
    
def StrDecrypt(infostr):
    '''
    功能：对输入的字符串解密，解密算法为先用base64方法解密，然后将字符串倒序，去掉最后的换行符，最后将字符串中的数字减去1
    '''
    infostr = base64.decodestring(infostr)[::-1].strip('\n')
    info = ''
    for i in range(len(infostr)):
        if infostr[i] >= '1' and infostr[i] <= '9':
            info += str(int(infostr[i]) - 1)
        elif infostr[i] == '0':
            info += '9'
        else:
            info += infostr[i]
    return info

def ProductRegedit(hardwareinfofile):
    '''
    功能：对硬件信息进行加密，并输出到lic文件，输出内容分别占一行，
    输出顺序为CPU、MAC、主板和硬盘
    '''
    
    #从文件中读取硬件信息
    if not os.path.exists(hardwareinfofile):
        return 0
    
    file = open(hardwareinfofile)
    #读取第一行，CPU信息
    cpuid = file.readline()
    #读取第二行，MAC信息
    macaddr = file.readline()
    #读取第三行，主板信息
    boardid = file.readline()
    #读取第四行，硬盘信息
    diskid = file.readline()
    file.close()
    
    #对上述信息进行加密
    allinfo = ''
    allinfo += StrEncrypt(cpuid)
    allinfo += StrEncrypt(macaddr)
    allinfo += StrEncrypt(boardid)
    allinfo += StrEncrypt(diskid)

    #写入文件
    ipos = hardwareinfofile.rfind('\\')
    if ipos > -1:
        licfilepath = hardwareinfofile[0:ipos+1] + 'xfdoc.lic'
    else:
        licfilepath = 'xfdoc.lic'
    f1 = open(licfilepath, 'w')#
    f1.writelines(allinfo)
    f1.close()
    return 1#成功生成注册文件
    
def ReadHardwareInfoFromRegeditFile(licfile):
    '''
    功能：从文件中读取加密后的硬件信息并进行解密
    '''
    #检查文件是否存在

    if not os.path.exists(licfile):
        return 0,'','','',''
    
    file = open(licfile)
    #读取第一行，CPU信息
    str = file.readline()
    
    cpuid = StrDecrypt(str)
    #读取第二行，MAC信息
    str = file.readline()
    macaddr = StrDecrypt(str)
    #读取第三行，主板信息
    str = file.readline()
    boardid = StrDecrypt(str)
    #读取第四行，硬盘信息
    str = file.readline()
    diskid = StrDecrypt(str)
    file.close()
    return 1,cpuid,macaddr,boardid,diskid
    
def ProductAccredit():
    '''
    功能：在软件启动开始调用，判断读取的硬件信息与加密后保存在本地的硬件信息是否一致
    返回值：
    1：认证成功
    0：认证失败
    -1：注册文件不存在
    '''
    cpuid,macaddr,boardid,diskid = GetHardwareInfo()
    
    #从文件中读取注册时提供的硬件信息
    flag,cpuidf,macaddrf,boardidf,diskidf = ReadHardwareInfoFromRegeditFile('xfdoc.lic')
    if flag == 1:
        #判断当前用户是否注册用户
        if macaddr.encode('utf-8') == macaddrf and cpuid.encode('utf-8') == cpuidf and \
                   boardid.encode('utf-8') == boardidf and diskid.encode('utf-8') == diskidf:
            return 1#软件认证成功
        else:
            return 0#软件认证失败
    else:
        return -1#注册文件不存在
        
#ProductRegedit('BFEBFBFF000206A7.txt')     
#print ProductAccredit()
#WriteHardwareInfotoFile()
#GetHardwareInfo()
