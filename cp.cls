VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub set_system()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.cu_system(0).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoConfigPage", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NoConfigPage", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_system(1).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoDevMgrPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoDevMgrPage", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_system(2).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoFileSysPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoFileSysPage", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_system(3).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoVirtMemPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoVirtMemPage", 0, 4, 0, 4)
       '*********** error1
    End If
End Sub

Public Sub set_network()
    Dim create_open As Long
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    '************* error
    
    If (Form1.cu_network(0).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoFileSharingControl", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoFileSharingControl", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_network(1).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoNetSetupIDPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoNetSetupIDPage", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_network(2).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoNetSetupSecurityPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoNetSetupSecurityPage", 0, 4, 0, 4)
       '*********** error1
    End If
End Sub

Public Sub set_passwords()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.cu_passwords(0).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoAdminPage", 0, 4, 1, 4)
    Else
        retval = RegSetValueEx(opened, "NoAdminPage", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_passwords(1).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoProfilePage", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NoProfilePage", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_passwords(2).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoPwdPage", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NoPwdPage", 0, 4, 0, 4)
        '*********** error1
    End If
    End Sub

Public Sub set_printers()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.cu_printers(0).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoAddPrinter", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoAddPrinter", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_printers(1).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoDeletePrinter", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoDeletePrinter", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_printers(2).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoPrinterTabs", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoPrinterTabs", 0, 4, 0, 4)
       '*********** error1
    End If
End Sub

Public Sub set_display()
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    If (Form1.cu_display(0).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoDispAppearancePage", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NoDispAppearancePage", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_display(1).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoDispBackgroundPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoDispBackgroundPage", 0, 4, 0, 4)
       '*********** error1
    End If
    
    If (Form1.cu_display(2).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoDispScrSavPage", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NoDispScrSavPage", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_display(3).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoDispSettingsPage", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoDispSettingsPage", 0, 4, 0, 4)
       '*********** error1
    End If
End Sub

Public Sub set_start_menu()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.cu_start_menu(1).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoRun", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NoRun", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_start_menu(2).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NosetTaskbar", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "NosetTaskbar", 0, 4, 0, 4)
        '*********** error1
    End If
    
    If (Form1.cu_start_menu(3).Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoFind", 0, 4, 1, 4)
       '*********** error1
    Else
       retval = RegSetValueEx(opened, "NoFind", 0, 4, 0, 4)
       '*********** error1
    End If

    If (Form1.cu_start_menu(4).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoRecentDocsHistory", 0, 4, 1, 4)
    Else
        retval = RegSetValueEx(opened, "NoRecentDocsHistory", 0, 4, 0, 4)
    End If
    
End Sub

Public Sub get_start_menu()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        
    retval = RegQueryValueEx(opened, "NoRun", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_start_menu(1).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NosetTaskbar", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_start_menu(2).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoFind", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_start_menu(3).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoRecentDocsHistory", 0, datatype, Lvalue_read, 4)
    If (Lvalue_read = 1) Then
        Form1.cu_start_menu(4).Value = vbChecked
        Lvalue_read = 0
    End If
    
End Sub

Public Sub get_system()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    retval = RegQueryValueEx(opened, "NoConfigPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_system(0).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoDevMgrPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_system(1).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoFileSysPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_system(2).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoVirtMemPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_system(3).Value = vbChecked
        Lvalue_read = 0
    End If
    
End Sub

Public Sub get_passwords()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    retval = RegQueryValueEx(opened, "NoAdminPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_passwords(0).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoProfilePage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_passwords(1).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoPwdPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_passwords(2).Value = vbChecked
        Lvalue_read = 0
    End If
End Sub

Public Sub get_network()
    
    Dim create_open As Long
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    '************* error
    retval = RegQueryValueEx(opened, "NoFileSharingControl", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_network(0).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoNetSetupIDPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_network(1).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoNetSetupSecurityPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_network(2).Value = vbChecked
        Lvalue_read = 0
    End If
    
End Sub

Public Sub get_display()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    retval = RegQueryValueEx(opened, "NoDispAppearancePage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_display(0).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoDispBackgroundPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_display(1).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoDispScrSavPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_display(2).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoDispSettingsPage", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_display(3).Value = vbChecked
        Lvalue_read = 0
    End If
    
End Sub

Public Sub get_printers()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    retval = RegQueryValueEx(opened, "NoAddPrinter", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_printers(0).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoDeletePrinter", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_printers(1).Value = vbChecked
        Lvalue_read = 0
    End If
    
    retval = RegQueryValueEx(opened, "NoPrinterTabs", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_printers(2).Value = vbChecked
        Lvalue_read = 0
    End If
End Sub



Public Sub get_cp()
   subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
   retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
   '************* error
   
   retval = RegQueryValueEx(opened, "NoControlPanel", 0, datatype, Lvalue_read, 4)
   '*********** error1
   
   If (Lvalue_read = 1) Then
      Form1.cu_cp.Value = vbChecked
      Lvalue_read = 0
   Else
      retval = RegQueryValueEx(opened, "NoSetFolders", 0, datatype, Lvalue_read, 4)
      '*********** error1
      
      If (Lvalue_read = 1) Then
         Form1.cu_cp.Value = vbChecked
         Lvalue_read = 0
      End If
   End If
End Sub

Public Sub set_cp()
   
   Dim create_open As Long
   subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
   
   retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
   '************* error
   
   If (Form1.cu_cp.Value = vbChecked) Then
       retval = RegSetValueEx(opened, "NoControlPanel", 0, 4, 1, 4)
       '*********** error1
       retval = RegSetValueEx(opened, "NoSetFolders", 0, 4, 1, 4)
       '*********** error1
       Lvalue_read = 0
   Else
       retval = RegSetValueEx(opened, "NoControlPanel", 0, 4, 0, 4)
       '*********** error1
       retval = RegSetValueEx(opened, "NoSetFolders", 0, 4, 0, 4)
       '*********** error1
       Lvalue_read = 0
    End If
End Sub

Public Sub get_disable()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"

   retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
   '************* error
   
   retval = RegQueryValueEx(opened, "DisableRegistryTools", 0, datatype, Lvalue_read, 4)
   '*********** error1
   
   If (Lvalue_read = 1) Then
      Form1.disable_registry.Value = vbChecked
      Lvalue_read = 0
   End If
End Sub

Public Sub set_disable()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.disable_registry.Value = vbChecked) Then
        retval = RegSetValueEx(opened, "DisableRegistryTools", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "DisableRegistryTools", 0, 4, 0, 4)
        '*********** error1
    End If
End Sub

