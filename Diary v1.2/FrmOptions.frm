VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   6795
   ClientTop       =   4365
   ClientWidth     =   1695
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   1695
   Begin VB.CheckBox Check1 
      Caption         =   "Always use Password startup."
      Height          =   375
      Left            =   120
      MouseIcon       =   "FrmOptions.frx":0442
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton CmdChangePass 
      Caption         =   "&Change Password"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '***************************************************************'
    '                       Diary V1.2.1                            '
    '                        Written by                             '
    '                         Cooliced                              '
    '                                                               '
    '  You are free to use the source code in your private,         '
    '  non-commercial, projects with permission.    If you want     '
    '  to use this code in commercial projects EXPLICIT permission  '
    '  from the author is required.                                 '
    '                                                               '
    '                                                               '
    '        Copyright © Cooliced - Cooliced.c.uk 1999-2000         '
    '***************************************************************'


Private Sub CmdChangePass_Click()
    FrmPassChange.Show
    Me.Hide
End Sub

Private Sub CmdOk_Click()
Dim BoxVal As String

If Check1.Value = 0 Then
 BoxVal = "no"
 PutValue "Startup", "Login", BoxVal, App.Path & "\" & con_INI_File
End If
If Check1.Value = 1 Then
 BoxVal = "yes"
 PutValue "Startup", "Login", BoxVal, App.Path & "\" & con_INI_File
End If

Unload Me
FrmMain.Show

End Sub

Private Sub Form_Load()

 Dim ChkVal As String
  ChkVal = GetValue("Startup", "Login", App.Path & "\" & con_INI_File)
    If ChkVal = "no" Then
    Check1.Value = 0
    End If
    If ChkVal = "yes" Then
    Check1.Value = 1
    End If
    
End Sub
