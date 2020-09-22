Attribute VB_Name = "cryptmod"
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


Public Function CryptString2(txtString As String, Encrypt As Boolean) As String


       On Error GoTo errhandler
       Dim x As Integer
       Dim outString As String
       Dim iLen As Integer
       Dim sFirstSeed As String
       Dim sSecondSeed As String
       Dim iSeed As Integer


       If Encrypt Then
           sFirstSeed = Left(txtString, 1)
           sSecondSeed = Mid(txtString, 2, 1)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString)


           For x = 1 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) + 2) & outString
           Next


           outString = Chr(Asc(sFirstSeed) * 2 + 3) & outString
           outString = outString & Chr(Asc(sSecondSeed) * 2 - 3)
       Else
           sFirstSeed = Chr((Asc(Left(txtString, 1)) - 3) \ 2)
           sSecondSeed = Chr((Asc(Right(txtString, 1)) + 3) \ 2)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString) - 1


           For x = 2 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) - 2) & outString
           Next


       End If


       CryptString2 = outString
       Exit Function
errhandler:
       MsgBox "Error in Diary" & vbCrLf & "Error: " & Err.Description & vbCrLf & "Number: " & Err.Number
       CryptString2 = ""
   End Function
