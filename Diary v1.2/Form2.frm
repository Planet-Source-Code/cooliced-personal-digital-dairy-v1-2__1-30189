VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Record"
   ClientHeight    =   5340
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4320
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Data"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancel Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Picture"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0472
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":08C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0D1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4560
      Top             =   480
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4965
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5777
            MinWidth        =   5645
            Text            =   "Caption"
            TextSave        =   "Caption"
            Object.ToolTipText     =   "Caption"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   50
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtBirth 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtSex 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   5
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtMobile 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox TxtPhoto 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Picture Link"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Telephone No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Mobile No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPicture 
         Caption         =   "&Insert Picture"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyData As Database
Dim MyRecord As Recordset

Private Sub cmdCancel_Click()
txtName.Text = ""
txtAddress.Text = ""
txtSex.Text = ""
txtBirth.Text = ""
txtEmail.Text = ""
txtPhone.Text = ""
txtMobile.Text = ""
TxtPhoto.Text = ""
Form1.Show
Form2.Hide
End Sub

Private Sub cmdSave_Click()
Set MyData = OpenDatabase(App.Path + "\Address.mdb")
Set MyRecord = MyData.OpenRecordset("AddressBook")

If txtName.Text = "" Then
    If MsgBox("Name Field is blank." & (Chr(10)) & "Please enter value.", vbInformation, "Enter Value") = vbOK Then
        'txtName.Text = "Nil"
        txtName.SetFocus
        Exit Sub
    End If
End If
Do While txtAddress.Text = "" Or txtBirth.Text = "" Or txtSex.Text = "" Or txtEmail.Text = "" _
        Or txtPhone.Text = "" Or txtMobile.Text = ""

If txtAddress.Text = "" Then
    If MsgBox("Address Field is blank." & (Chr(10)) & "Nil is enter.", vbInformation, "Enter Value") = vbOK Then
        txtAddress.Text = "Nil"
        'txtAddress.SetFocus
        'Exit Sub
    End If
ElseIf txtBirth.Text = "" Then
    If MsgBox("D.O.B Field is blank." & (Chr(10)) & "Nil is enter.", vbInformation, "Enter Value") = vbOK Then
        txtBirth.Text = "Nil"
        'txtBirth.SetFocus
        'Exit Sub
    End If
ElseIf txtSex.Text = "" Then
    If MsgBox("Gender Field is blank." & (Chr(10)) & "Nil is enter.", vbInformation, "Enter Value") = vbOK Then
        txtSex.Text = "Nil"
        'txtSex.SetFocus
        'Exit Sub
    End If
ElseIf txtEmail.Text = "" Then
    If MsgBox("E-Mail Field is blank." & (Chr(10)) & "Nil is enter.", vbInformation, "Enter Value") = vbOK Then
        txtEmail.Text = "Nil"
        'txtEmail.SetFocus
        'Exit Sub
    End If
ElseIf txtPhone.Text = "" Then
    If MsgBox("Telephone No. Field is blank." & (Chr(10)) & "Nil is enter.", vbInformation, "Enter Value") = vbOK Then
        txtPhone.Text = "Nil"
        'txtPhone.SetFocus
        'Exit Sub
    End If
ElseIf txtMobile.Text = "" Then
    If MsgBox("Mobile No. Field is blank." & (Chr(10)) & "Nil is enter.", vbInformation, "Enter Value") = vbOK Then
        txtMobile.Text = "Nil"
        'txtMobile.SetFocus
        'Exit Sub
    End If
End If
Loop

With MyRecord
        .AddNew
        !Name = Trim(txtName.Text)
        !Address = Trim(txtAddress.Text)
        !Birth_Date = Trim(txtBirth.Text)
        !Sex = Trim(txtSex.Text)
        !Email_Address = Trim(txtEmail.Text)
        !Home_Phone = Trim(txtPhone.Text)
        !MobilePhone = Trim(txtMobile.Text)
        !Photo = Trim(TxtPhoto.Text)
        .Update
    End With
Form1.List1.AddItem txtName.Text
Form1.List1.ListIndex = 0
txtName.Text = ""
txtAddress.Text = ""
txtSex.Text = ""
txtBirth.Text = ""
txtEmail.Text = ""
txtPhone.Text = ""
txtMobile.Text = ""
TxtPhoto.Text = ""
Form1.Show
Form2.Hide

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call cmdCancel_Click
End Sub


Private Sub mnuCancel_Click()
Call cmdCancel_Click
End Sub

Private Sub mnuPicture_Click()
On Error GoTo DialogError
With CommonDialog1
        .CancelError = True
        .Filter = "JPG File (*.jpg)|*.jpg|Bitmap File (*.bmp)|*.bmp|GIF File(*.gif)|*.gif|All Files(*.*)|*.*"
        .FilterIndex = 1
        .DialogTitle = "Select a Picture File"
        .ShowOpen
   TxtPhoto.Text = .FileName
   
   End With
DialogError:
End Sub

Private Sub mnuSave_Click()
Call cmdSave_Click
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels.Item(1) = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call cmdSave_Click
    Case 2
        Call cmdCancel_Click
    Case 3
        Call mnuPicture_Click
   
End Select
    
End Sub

Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Address Field"
End Sub

Private Sub txtBirth_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Date of Birth Field"
End Sub

Private Sub txtEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Email Address Field "
End Sub

Private Sub txtMobile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Mobile No. Field"
End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Name Field"
End Sub

Private Sub txtPhone_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Home No. Field"
End Sub

Private Sub TxtPhoto_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Picture Link Field"
End Sub

Private Sub txtSex_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "Gender Field"
End Sub

