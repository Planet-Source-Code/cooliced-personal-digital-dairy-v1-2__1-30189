VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2580
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   4680
      Y1              =   2160
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   630
      X2              =   5280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Getting Settings...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   2295
      Width           =   3435
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   465
      Left            =   630
      Top             =   2160
      Width           =   4380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00FF0000&
      Height          =   2670
      Left            =   -45
      Top             =   -45
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Diary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   1350
      TabIndex        =   5
      Top             =   -45
      Width           =   2445
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1350
      TabIndex        =   2
      Top             =   765
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cooliced"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   945
      Width           =   2445
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyrights (c) 2001 Cooliced, All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   4020
   End
End
Attribute VB_Name = "frmSplash"
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
Option Explicit

Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
    lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
    lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
    
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
Dim strString As String
    Dim lngDword As Long


    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Sub Startup()
Dim ChkVal As String
Dim BoxVal As String
Dim strText As String
Dim PassVal As String

    ' Show form
    frmSplash.Show
    ' Refresh form
    frmSplash.Refresh
    ' Call Loading sub
    Loading
    ' Close form
    Unload Me
    
    'Login startup
    If FileExists(con_INI_File) = False Then
        PassVal = ""
        PutValue "Main", "Password", PassVal, App.Path & "\" & con_INI_File
        
        BoxVal = "yes"
        PutValue "Startup", "Login", BoxVal, App.Path & "\" & con_INI_File
    End If
    If FileExists(con_INI_File) = True Then
    ChkVal = GetValue("Startup", "Login", App.Path & "\" & con_INI_File)
    ' ChkVal is the string it checks for in the ini file
     If ChkVal = "no" Then     ' if ChkVal = no then
        FrmMain.Show ' show the main form
     End If
     
     If ChkVal = "yes" Then    'if ChkVal = yes then
        FrmLogin.Show           'show login form
     End If
    End If
End Sub

Public Sub Loading()
      CStatus "Getting settings..."
    GetSettings 'Load settings from ini
End Sub

Public Sub CStatus(Message As String)
    ' Set label caption
    lblStatus.Caption = Message
    ' Refresh label
    lblStatus.Refresh
End Sub
