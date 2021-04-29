from oletools.olevba import VBA_Parser,TYPE_OLE,TYPE_OpenXML,TYPE_Word2003_XML,TYPE_MHTML 
import os
import time


badAPIlist = {

    'CreateMutex':0,
    'OpenMutex':0,
    'CreateFile':0,
    'CreateFileMapping':0,
    'WinExec':0,
    'CreateService':0,
    'CreateProcess':0,
    'ShellExecute':0,
    'SetWindowsHookEx':0,
    'IsDebuggerPresent':0,  
    'Sleep':0,
    'VirtualProtect':0,
    'URLDownloadToFile':0,
    'DeleteFile':0,
    'ReadProcessMemory':0,   
    'VirtualAlloc':0,   
    'VAlloc':0,
    'LoadLibraryA':0,   
    'GetCurrentProcess':0,
    'CreateThread':0,   
    'CreateRemoteThread':0,  
    'Socket':0,   
    'Send':0,   
    'Recv':0,   
    'Keybd_event':0,   
    'WriteFile':0,   
    'ReadFile':0,   
    'OpenProcess':0,   
    'WriteProcessMemory':0,   
    'FtpOpenFile':0,   
    'FtpPutFile':0,   
    'GetAsyncKeyState':0,
    'GetDiskFreeSpaceExW':0,   
    'CreateProcessA':0,   
    'InternetWriteFile':0,   
    'InternetConnect':0,   
    'Connect':0,    
    'malloc':0,   
    'AdjustTokenPrivileges':0,   
    'AttachThreadinput':0,   
    'Bind':0,   
    'FindResource':0,   
    'FindWindow':0,   
    'Gethostbyname':0,
    'InternetOpen':0,   
    'InternetOpenUrl':0,
    'InternetReadFile':0,   
    'NtQueryDirectoryFile':0,   
    'WSAStartup':0,     
    'MapVirtualKey':0,   
    'BitBlt':0,   
    'IsNTAdmin':0, 
    'RegOpenKey':0,   
    '.exe':0,
    '.vbs':0,   
    '.bat':0,
    'powershell':0,
    'del.exe /':0, 
    'Window':0,
    'Kerne132':0,   
    'del /':0,   
    'cmd /':0,    
    'taskmgr':0,   
    'taskkill':0,     
    'mshta':0,      
    'rundll':0,      
    'VboxHook':0,      
    'ollydbg':0 ,     
    'Vmware':0,      
    'Kernel32':0,      
    'system32':0 ,     
    'nc.exe':0       
}

class VBAChecker:
    
    badAPI = {
        'CreateMutex',
        'OpenMutex',
        'CreateFile',
        'CreateFileMapping',
        'WinExec',
        'CreateService',
        'CreateProcess',
        'ShellExecute',
        'SetWindowsHookEx',
        'IsDebuggerPresent',  
        'Sleep',
        'VirtualProtect',
        'URLDownloadToFile',
        'DeleteFile',
        'ReadProcessMemory',   
        'VirtualAlloc',   
        'VAlloc',
        'LoadLibraryA',   
        'GetCurrentProcess',
        'CreateThread',   
        'CreateRemoteThread',  
        'Socket',   
        'Send',   
        'Recv',   
        'Keybd_event',   
        'WriteFile',   
        'ReadFile',   
        'OpenProcess',   
        'WriteProcessMemory',   
        'FtpOpenFile',   
        'FtpPutFile',   
        'GetAsyncKeyState',
        'GetDiskFreeSpaceExW',   
        'CreateProcessA',   
        'InternetWriteFile',   
        'InternetConnect',   
        'Connect',    
        'malloc',   
        'AdjustTokenPrivileges',   
        'AttachThreadinput',     
        'Bind',   
        'FindResource',   
        'FindWindow',   
        'Gethostbyname',
        'InternetOpen',   
        'InternetOpenUrl',
        'InternetReadFile',   
        'NtQueryDirectoryFile',   
        'WSAStartup',   
        'IsNTAdmin',      
        'BitBlt', 
        'MapVirtualKey',   
        'RegOpenKey',   
        '.exe',
        '.vbs',   
        '.bat',
        'powershell',
        'del.exe /',
        'Window',
        'Kerne132',   
        'del /',   
        'cmd /',    
        'taskmgr',   
        'taskkill',   
        'mshta',  
        'rundll',      
        'VboxHook',      
        'ollydbg',    
        'Vmware',      
        'Kernel32',      
        'system32',      
        'nc.exe' 
    }

    def __init__(self,filep = None):
        if filep == None :
            self.filePath = ''
        else :
            self.filePath = filep

    def VCdetector(self):
        if self.filePath == '':
            return -1
        vbaparser = VBA_Parser(self.filePath)
        if vbaparser.detect_vba_macros():
            print(vbaparser.reveal())
            result= self.BadAPIchecker(vbaparser.reveal())
            vbaparser.close()
            return result
        else :
            print('macro not found')
            vbaparser.close()
            return False

    def BadAPIchecker(self, VC) :
        for each in self.badAPI :
            badAPIlist[each]=VC.count(each)
            
        print(badAPIlist['URLDownloadToFile'])
        return True ,badAPIlist


if __name__ == '__main__':
    start =time.time()
    result=VBAChecker('Dlib.docx').VCdetector()
    if result == False :
        print("Macro 가 없는 문서입니다.") #  안전파일 반환
    else :
        target = list(result[1].values()) #인공지능에 들어갈 배열
    print("time: ",time.time() - start)
