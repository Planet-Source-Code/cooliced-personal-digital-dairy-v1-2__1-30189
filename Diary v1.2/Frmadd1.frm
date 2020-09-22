VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   6915
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   8595
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewPic 
      Caption         =   "View Pic"
      Height          =   495
      Left            =   4680
      TabIndex        =   24
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdSort2 
      Caption         =   "Z TO A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   33
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Search List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   6720
      TabIndex        =   27
      Top             =   600
      Width           =   1695
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
      TabIndex        =   25
      Top             =   4440
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1448
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":283C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Data"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit Data"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove Data"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Email"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   720
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   6540
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10345
            Text            =   "Click Edit Button to edit data "
            TextSave        =   "Click Edit Button to edit data "
            Object.ToolTipText     =   "Caption"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2593
            MinWidth        =   2593
            Picture         =   "Form1.frx":3538
            Text            =   "Version 1.1"
            TextSave        =   "Version 1.1"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "A TO Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "Form1.frx":398C
      Left            =   4560
      List            =   "Form1.frx":398E
      TabIndex        =   15
      Top             =   600
      Width           =   1815
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
      TabIndex        =   14
      Top             =   3840
      Width           =   2295
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
      TabIndex        =   13
      Top             =   3240
      Width           =   2295
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data List"
      Height          =   2415
      Left            =   4440
      TabIndex        =   20
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Data"
      Height          =   495
      Left            =   3840
      TabIndex        =   17
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Data"
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Search"
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   5880
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search List"
      Height          =   4575
      Left            =   6600
      TabIndex        =   29
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Left            =   1560
      TabIndex        =   31
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Forward"
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   735
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
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Picture Link "
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
      TabIndex        =   26
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   4440
      MouseIcon       =   "Form1.frx":3990
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2055
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
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
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label9 
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
      TabIndex        =   30
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   4440
      TabIndex        =   32
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Data"
      End
      Begin VB.Menu mnuAddData 
         Caption         =   "&Add Data"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove Data"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "E&mail"
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "&Tool"
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuDescending 
         Caption         =   "Sort &Descending"
      End
      Begin VB.Menu mnuAscending 
         Caption         =   "Sort &Ascending"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim MyData As Database
Dim MyRecord As Recordset
Dim SQL As String

Private Sub cmdAdd_Click()
    Form2.Show
    Form1.Hide
    Form2.txtName.SetFocus
End Sub

Private Sub cmdBack_Click()
    If List1.ListIndex = 0 Then
        List1.ListIndex = (List1.ListCount - 1)
    Else
        List1.ListIndex = (List1.ListIndex - 1)
    End If
End Sub

Private Sub cmdClear_Click()
    List2.Clear
    Form1.Width = 6675
    Call cmdFirst_Click
End Sub

Private Sub cmdEdit_Click()

End Sub

Private Sub cmdFirst_Click()
If List1.ListCount = 1 Then
    'nothing
Else
    List1.ListIndex = 0
    List1.ListIndex = 1
End If
End Sub

Private Sub cmdInsert_Click()

End Sub

Private Sub cmdSort2_Click()
Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = List1.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With List1
        If .List(i) < .List(i + 1) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j
Call cmdFirst_Click
End Sub

Private Sub cmdSort2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Descending Order Z to A"
End Sub

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

Private Sub cmdfind_Click()
Set MyData = OpenDatabase(App.Path + "\Address.mdb")
Dim LName As String
Dim strFind As String
List2.Clear
strFind = Trim(InputBox("Enter Name for search.", "Search Box"))
LName = Trim(UCase(strFind))
If strFind = "" Then
    Exit Sub
End If
Set MyRecord = MyData.OpenRecordset("SELECT Name " & "FROM AddressBook " & _
                                    "WHERE Name Like '*" & LName & "*'")

With MyRecord
    If .EOF Then
        MsgBox "No matching Name found, try again please", vbCritical, "Result"
        Form1.Width = 6675
    Else
        Do Until .EOF
            Form1.Width = 8645
            List2.AddItem !Name
            .MoveNext
        Loop
    End If
End With

End Sub

Private Sub cmdNext_Click()
If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
Else
    List1.ListIndex = (List1.ListIndex + 1)
End If

End Sub

Private Sub cmdRemove_Click()
If List1.Text = "(Creator)" Then
    MsgBox "You cannot delete this data", vbExclamation, "Note"
Else
    Set MyData = OpenDatabase(App.Path + "\Address.mdb")
    SQL = "SELECT * FROM AddressBook"
    Set MyRecord = MyData.OpenRecordset(SQL)
    Do Until MyRecord.EOF
        If List1.Text = MyRecord!Name Then
            If MsgBox("You Really want to delete this record ?", vbQuestion + vbYesNo, "Delete Record") = vbYes Then
                MyRecord.Delete
                List1.RemoveItem (List1.ListIndex)
            End If
        End If
        MyRecord.MoveNext
    Loop
    Call cmdNext_Click
End If
End Sub

Private Sub cmdSort_Click()

Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = List1.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With List1
        If .List(i) > .List(i + 1) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j
Call cmdFirst_Click
End Sub

Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Ascending Order A to Z"
End Sub

Private Sub Form_Load()
Form1.Width = 6675
Form1.Height = 6545
Set MyData = OpenDatabase(App.Path + "\Address.mdb")
Set MyRecord = MyData.OpenRecordset("AddressBook")
    If MyRecord.EOF Then
        MsgBox "No Data Found In AddressBook", vbInformation, "Notice"
    Else
        MyRecord.MoveFirst
        Do Until MyRecord.EOF
            List1.AddItem MyRecord.Fields("Name")
            MyRecord.MoveNext
        Loop
        Call cmdNext_Click
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Click Edit Button to edit data "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Please Come Again, Bye!!", vbInformation, "Exit"
    End
End Sub

Private Sub LabMailMe_DblClick()
Dim MyVal As Long
    MyVal = Shell("start mailto:" + "zijiang11@hotmail.com", 0)
End Sub

Private Sub List1_Click()
On Error Resume Next
        Set MyData = OpenDatabase(App.Path + "\Address.mdb")
        Set MyRecord = MyData.OpenRecordset("AddressBook")
        MyRecord.MoveFirst
    Do Until MyRecord.EOF
            If List1.Text = MyRecord!Name Then
                txtName.Text = MyRecord!Name
                txtAddress.Text = MyRecord!Address
                txtSex.Text = MyRecord!Sex
                txtBirth.Text = MyRecord!Birth_Date
                txtEmail.Text = MyRecord!Email_Address
                txtPhone.Text = MyRecord!Home_Phone
                txtMobile.Text = MyRecord!MobilePhone
                TxtPhoto.Text = MyRecord!Photo
            End If
            MyRecord.MoveNext
    Loop

If TxtPhoto.Text = "" Then
    Image1.Picture = LoadPicture("")
Else
    Image1.Picture = LoadPicture("")
    Call cmdViewPic_Click
End If
mnuEdit.Enabled = True
Toolbar1.Buttons(4).Enabled = True
End Sub

Private Sub List2_Click()
On Error Resume Next
        Set MyData = OpenDatabase(App.Path + "\Address.mdb")
        Set MyRecord = MyData.OpenRecordset("AddressBook")
        MyRecord.MoveFirst
    Do Until MyRecord.EOF
            If List2.Text = MyRecord!Name Then
                txtName.Text = MyRecord!Name
                txtAddress.Text = MyRecord!Address
                txtSex.Text = MyRecord!Sex
                txtBirth.Text = MyRecord!Birth_Date
                txtEmail.Text = MyRecord!Email_Address
                txtPhone.Text = MyRecord!Home_Phone
                txtMobile.Text = MyRecord!MobilePhone
                TxtPhoto.Text = MyRecord!Photo
            End If
            MyRecord.MoveNext
    Loop

If TxtPhoto.Text = "" Then
    Image1.Picture = LoadPicture("")
Else
    Image1.Picture = LoadPicture("")
    Call cmdViewPic_Click
End If
mnuEdit.Enabled = False
Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub mnuAbout_Click()
Dim iReturnValue
    iReturnValue = Shell("Notepad " & (App.Path + "\Readme file.txt"), 1)
End Sub

Private Sub mnuAddData_Click()
    Call cmdAdd_Click
End Sub

Private Sub mnuAscending_Click()
    Call cmdSort_Click
End Sub

Private Sub mnuDescending_Click()
    Call cmdSort2_Click
End Sub

Private Sub mnuEmail_Click()
Dim RetVal As Long
If txtEmail.Text = "nil" Or txtEmail.Text = "na" Then
    MsgBox "Not a valid Address", vbInformation, "Info"
Else
    RetVal = Shell("start mailto:" + (txtEmail.Text), 0)
End If
End Sub

Private Sub mnuExit_Click()
    If MsgBox("You really want to Exit ?", vbQuestion + vbYesNo, "Exit Address Book") = vbYes Then
        End
    End If
End Sub

Private Sub mnuHTML_Click()
With Form4
    .txtPath = App.Path
    .Show 1
End With
End Sub

Private Sub mnuRemove_Click()
    Call cmdRemove_Click
End Sub

Private Sub mnuEdit_Click()
If List1.Text = "(Creator)" Then
    MsgBox "You cannot edit this data", vbExclamation, "Note"
Else
    Form3.Show
    Form1.Hide
End If
End Sub

Private Sub mnuSearch_Click()
    Call cmdfind_Click
End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels.Item(1) = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call cmdBack_Click
    Case 2
        Call cmdNext_Click
    Case 3
        Call cmdAdd_Click
    Case 4
        Call mnuEdit_Click
    Case 5
        Call cmdRemove_Click
    Case 6
        Call cmdfind_Click
    Case 7
        Call mnuEmail_Click
    Case 8
        Call mnuAbout_Click
    Case 9
        Call mnuExit_Click
End Select
        
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Address Field"
End Sub

Private Sub txtBirth_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtBirth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Date of Birth Field"
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Email Address Field "
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtMobile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Mobile No. Field"
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Name Field"
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtPhone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Home No. Field"
End Sub

Private Sub TxtPhoto_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub TxtPhoto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Picture Link Field"
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then KeyAscii = 0
If KeyAscii < Asc("0") Or KeyAscii > ("9") Then KeyAscii = 0
End Sub

Private Sub txtSex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(2).Text = "Gender Field"
End Sub
