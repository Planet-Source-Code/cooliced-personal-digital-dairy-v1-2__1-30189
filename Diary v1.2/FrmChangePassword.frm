VERSION 5.00
Begin VB.Form FrmPassChange 
   Caption         =   "Change Password"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   Icon            =   "FrmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame FrChangePassword 
      Caption         =   "Change Password"
      Height          =   2415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtExistingPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Existing Password"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Enter New Password"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm New Password"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmPassChange"
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


Private Sub cmdCancel_Click()
    Unload Me
    FrmOptions.Show
End Sub


Private Sub CmdOk_Click()

Dim strTemp As String
Dim strPW As String
Dim strNewPW As String
Dim strEncryptNewPW As String
    'some error handling
    
    strPW = GetValue("Main", "Password", App.Path & "\" & con_INI_File)
    strNewPW = LCase(txtNewPassword2.Text)
    'checks to see if you type in the correct password in the existing password field
        
     If FrmLogin.txtPassword = LCase(txtExistingPassword.Text) Then
        'checks the match of the new passwords
        
        If LCase(txtNewPassword1.Text) = strNewPW Then
            strEncryptNewPW = Encrypt(strNewPW)
            PutValue "Main", "Password", strEncryptNewPW, App.Path & "\" & con_INI_File
            MsgBox "Password changed!", 8, "Password Verfication"
        
        Else
            MsgBox "The New Passwords Do Not Match", 8, "Password Error"
            txtNewPassword1.SetFocus
            Exit Sub
        
        End If
        
    Else
        MsgBox "The Existing Password is Incorrect!", 8, "Password Error"
        txtExistingPassword.SetFocus
        Exit Sub
        
    End If
    
    FrmPassChange.Hide
    DoEvents
    
End Sub

