VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmText 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5055
   ClientLeft      =   5415
   ClientTop       =   2985
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393217
      TextRTF         =   $"FrmText.frx":0000
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8916
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"FrmText.frx":00AE
   End
   Begin VB.Menu DoSave 
      Caption         =   "S&ave"
   End
   Begin VB.Menu DoClose 
      Caption         =   "C&lose"
   End
End
Attribute VB_Name = "FrmText"
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


Private Sub DoClose_Click()
 FrmMain.Show
 Unload Me
End Sub

Private Sub DoSave_Click()

     Path = App.Path & "\Data\"
     extention = ".ccd"
     FileName = Format(FrmMain.Calendar1.Value, "Medium Date")

    If RTB1.Text = "" Then
     DoSave.Checked = False
      Exit Sub
     Else
      RTB2.Text = CryptString2(RTB1.Text, True)
      RTB2.SaveFile Path & FileName & extention
     MsgBox "Saved", vbInformation Or vbOKOnly, "Saved"
     End If
     
End Sub

Private Sub Form_Load()
 Me.Caption = Format(FrmMain.Calendar1.Value, "Long Date")
 
 Path = App.Path & "\Data\"
 extention = ".ccd"
 FileName = Format(FrmMain.Calendar1.Value, "Medium Date")

 If FileExists(Path & FileName & extention) = False Then
  MsgBox "This date is empty at the moment", vbExclamation Or vbOKOnly, "Diary Data"
  RTB1.Text = ""
 Else
  RTB2.LoadFile Path & FileName & extention
  RTB1.Text = CryptString2(RTB2.Text, False)
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

 If RTB1.Text = "" Then
  Unload Me
 Else
  FrmMain.Show
  Unload Me
 End If
  FrmMain.Show
 End Sub
