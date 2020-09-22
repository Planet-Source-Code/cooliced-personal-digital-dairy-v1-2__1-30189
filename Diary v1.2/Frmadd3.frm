VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Edit Record"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   6210
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewPic 
      Caption         =   "View Pic"
      Height          =   495
      Left            =   6360
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Edit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancel Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit Picture"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove Picture"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":08BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":12AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   1200
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
      TabIndex        =   8
      Top             =   4440
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
      TabIndex        =   7
      Top             =   3840
      Width           =   2535
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
      TabIndex        =   6
      Top             =   3240
      Width           =   2535
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
      TabIndex        =   5
      Top             =   2640
      Width           =   3255
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
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
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
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
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
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
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
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4935
      Width           =   6210
      _ExtentX        =   10954
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
            Object.Width           =   8661
            MinWidth        =   5645
            Text            =   "Caption"
            TextSave        =   "Caption"
            Object.ToolTipText     =   "Caption"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   4320
      MouseIcon       =   "Form3.frx":16FE
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1815
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
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
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
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
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
      TabIndex        =   14
      Top             =   2640
      Width           =   735
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
      TabIndex        =   12
      Top             =   1920
      Width           =   975
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
      TabIndex        =   11
      Top             =   1080
      Width           =   975
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
      TabIndex        =   10
      Top             =   600
      Width           =   735
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
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "No Picture Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   4560
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Double Click"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   4320
      TabIndex        =   18
      Top             =   720
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPicture 
         Caption         =   "&Edit Picture"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove Picture"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Save Edit"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel Edit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyData As Database
Dim MyRecord As Recordset

Private Sub cmdViewPic_Click()
Dim sExtension As String
 sExtension = UCase(Right$(TxtPhoto, 3))
 
 If Dir$(TxtPhoto) = "" Then
    MsgBox "Invaild Path", vbExclamation, "Picture Link Field "
    Exit Sub
 End If
 
 Select Case sExtension
    Case "JPG", "GIF", "BMP"
        Image1.Picture = LoadPicture(TxtPhoto.Text)
    Case Else
        MsgBox "Invaild Path", vbExclamation, "Link Field"
End Select
End Sub

Private Sub Form_Load()
txtName.Text = Form1.txtName.Text
txtAddress.Text = Form1.txtAddress.Text
txtSex.Text = Form1.txtSex.Text
txtBirth.Text = Form1.txtBirth.Text
txtEmail.Text = Form1.txtEmail.Text
txtPhone.Text = Form1.txtPhone.Text
txtMobile.Text = Form1.txtMobile.Text
TxtPhoto.Text = Form1.TxtPhoto.Text
If TxtPhoto.Text = "" Then
    Image1.Picture = LoadPicture("")
Else
    Image1.Picture = LoadPicture("")
    Call cmdViewPic_Click
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mnuCancel_Click
End Sub

Private Sub Image1_DblClick()
    Call mnuPicture_Click
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Double click to change picture"
End Sub

Private Sub mnuAbout_Click()
Dim iReturnValue
    iReturnValue = Shell("Notepad " & (App.Path + "\Readme file.txt"), 1)
End Sub

Private Sub mnuCancel_Click()
txtName.Text = ""
txtAddress.Text = ""
txtSex.Text = ""
txtBirth.Text = ""
txtEmail.Text = ""
txtPhone.Text = ""
txtMobile.Text = ""
TxtPhoto.Text = ""
Form1.Show
Unload Form3
End Sub

Private Sub mnuEdit_Click()
Set MyData = OpenDatabase(App.Path + "\Address.mdb")
Set MyRecord = MyData.OpenRecordset("AddressBook")

If txtName.Text = "" Then
    If MsgBox("Name Field is blank." & (Chr(10)) & "Please enter value.", vbInformation, "Enter Value") = vbOK Then
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
    
    Do Until MyRecord.EOF
        If Form1.List1.Text = MyRecord!Name Then
            Form1.List1.RemoveItem (Form1.List1.ListIndex)
            Form1.List1.AddItem txtName.Text
            With MyRecord
                .Edit
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
        End If
        MyRecord.MoveNext
    Loop
Form1.Show
Form1.List1.ListIndex = 1
Form1.List1.ListIndex = 0
Unload Form3
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
Call cmdViewPic_Click
End Sub

Private Sub mnuRemove_Click()
Dim Cutpic As String
If TxtPhoto.Text = "" Then
    'Do nothing
Else
    Cutpic = MsgBox("You want to erase the image path ?", vbQuestion + vbYesNo, "Erase Picture")
        If Cutpic = vbYes Then
            TxtPhoto.Text = ""
            Image1.Picture = LoadPicture("")
        End If
End If
End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels.Item(1) = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call mnuEdit_Click
    Case 2
        Call mnuCancel_Click
    Case 3
        Call mnuPicture_Click
    Case 4
        Call mnuRemove_Click
    Case 5
        Call mnuAbout_Click
End Select
End Sub
Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Address Field"
End Sub

Private Sub txtBirth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Date of Birth Field"
End Sub

Private Sub txtEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Email Address Field "
End Sub

Private Sub txtMobile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Mobile No. Field"
End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Name Field"
End Sub

Private Sub txtPhone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Home No. Field"
End Sub

Private Sub TxtPhoto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Picture Link Field"
End Sub

Private Sub txtSex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Gender Field"
End Sub
