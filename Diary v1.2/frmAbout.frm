VERSION 5.00
Begin VB.Form FrmAbout 
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmAbout"
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


Private Sub Form_Unload(Cancel As Integer)

FrmMain.Show
Unload Me
End Sub
