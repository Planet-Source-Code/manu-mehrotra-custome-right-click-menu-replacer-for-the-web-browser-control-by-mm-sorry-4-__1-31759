VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Load Main Form"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "O.K"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmd_apply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox ie_options 
      Caption         =   "Please Check If Using For First Time"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_apply_Click()
Call set_ie
End Sub

Private Sub cmd_ok_Click()



Call set_ie


'********** DETERMINING ENABLING PASSWORD


If (vbOK = MsgBox("Some of the settings may require the system to reboot.Do you want to restart now?", vbInformation + vbOKCancel, "Confirmation")) Then
    Call ExitWindowsEx(2, 0)
End If
End
End Sub


Private Sub Command1_Click()
Form11.Hide
Form1.Show

End Sub

Private Sub Form_Load()
 
    
    
Call get_ie
End Sub


Private Sub ie_options_Click()
    cmd_apply.Enabled = True
End Sub

