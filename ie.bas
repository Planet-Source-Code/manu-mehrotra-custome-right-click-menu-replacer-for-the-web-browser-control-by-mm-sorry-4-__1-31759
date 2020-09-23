Attribute VB_Name = "ie"
Option Explicit

Public Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Boolean
End Type

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const HKEY_CURRENT_USER = &H80000001
Public state_read As String * 3
Public retval As Long
Public datatype As Long
Public Lvalue_read As Long
Public Svalue_read As String * 255
Public opened As Long
Public secattr As SECURITY_ATTRIBUTES
Public subkey As String
Public Sub get_ie()
    Dim create_open As Long
   
 
 
    
        
    
    
    '**************** GETTING THE IE CAPION
     
        
    
    
        
    '
    
    subkey = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    
    '
    
    '************** GETTING THE RIGHT CLICK MENU ****************
           subkey = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
        If (retval = 0 And Lvalue_read = 1) Then
            Form11.ie_options.Value = vbChecked
        End If
        
    
   

End Sub

Public Sub set_ie()
    
    Dim create_open As Long
    Dim temp_string As String
    
   
    
   
    '************** SETTING THE RIGHT CLICK MENU ****************
        If (Form11.ie_options.Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoBrowserContextMenu", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoBrowserContextMenu", 0, 4, 0, 4)
        End If
    
   
            
            
End Sub

