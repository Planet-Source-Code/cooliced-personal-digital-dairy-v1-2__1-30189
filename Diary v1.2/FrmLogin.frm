VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Protected"
   ClientHeight    =   1245
   ClientLeft      =   5100
   ClientTop       =   5325
   ClientWidth     =   5070
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5070
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&OK"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Password"
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   480
         MouseIcon       =   "FrmLogin.frx":0442
         Picture         =   "FrmLogin.frx":074C
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmLogin"
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
    End
End Sub

Private Sub cmdSubmit_Click()

    Dim strTest As String
    strTest = GetValue("Main", "Password", App.Path & "\" & con_INI_File) ' get password
   
     If LCase(txtPassword.Text) = Decrypt(strTest) Then ' if textbox text = ini file decrypted text
        ' show
        FrmMain.Show
        ' The name of the main application
        Me.Hide
        ' Hides the login dialog box
        
    Else 'incorrect password!
        MsgBox "Enter a Valid Password for this System", 8, "Password Error"
        txtPassword.SetFocus
        Exit Sub
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End ' end program
End Sub
