VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public gsempnum As String
Public gwksED As Worksheet, gwksCL As Worksheet, gwksML As Worksheet, gwksRC As Worksheet, giX As Integer, giI As Integer, gsCourtOS As String
Public giY As Integer, gsDismq As String, giCo As Integer, gdteLeaveDSAP As Date, gsLeaveDSAP As String, giLSyear As Integer, giLSmonth As Integer, giLSday As Integer
Public gsLeavingR As String, gsLeaveD As String, gdteLeaveDD As Date, gsLastDM As String, giLyear As Integer, giLmonth As Integer, giLday As Integer
Public grngCl As Range, gsAbsLeave As String, gdteAbsLeave As Date, giALyear As Integer, giALmonth As Integer, giALday As Integer, gs655D As String, gdte655D As Date
Public gdLeaveD As Date, wsUserData As Worksheet

Public SapGuiAuto
Public SAPApp
Public SAPCon
Public session

Public gsLeavingC As String

Public c As Integer, gsempnumH As String, grEmpNumRng As Range, cell As Range, WS As Worksheet

Private Sub CommandButton1_Click()

UserForm10.Show













'Dim WB As Workbook
'Dim WSLF As Worksheet, WSML As Worksheet
'Dim sfindVal As String, LeaveD As String, LeavingM As String, LeavingY As String, Months(1 To 12) As String
'Dim cl As Range, EmpNum As String, EmpName As String, EmpFchar As String
'
'Set WB = Application.ThisWorkbook
'Set WSLF = Application.ThisWorkbook.Worksheets("Leaver Folder")
'Set WSML = ThisWorkbook.Worksheets("Managing Leavers")
'
'sfindVal = ActiveCell.Value
'EmpNum = sfindVal
'
'With WSML
'Set cl = .Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
'If Not cl Is Nothing Then cl.Select
'End With
'
'EmpName = WSML.Cells(cl.Row, "c")
'EmpFchar = Left(EmpName, 1)
'LeaveD = WSML.Cells(cl.Row, "e")
'LeavingM = Split(LeaveD, "/")(1)
'LeavingY = Split(LeaveD, "/")(2)
'
'Months(1) = "January"
'Months(2) = "February"
'Months(3) = "March"
'Months(4) = "April"
'Months(5) = "May"
'Months(6) = "June"
'Months(7) = "July"
'Months(8) = "August"
'Months(9) = "September"
'Months(10) = "October"
'Months(11) = "November"
'Months(12) = "December"
'
'    With WB.Connections("Connection").OLEDBConnection
'        .BackgroundQuery = False
'        .CommandText = Array( _
'        "<LIST><VIEWGUID>{42EEE4CF-14EF-4B58-B95D-61924126CBE2}</VIEWGUID><LISTNAME>{B8910762-E3D1-4AA2-A04E-7D516CED1381}</" _
'        , _
'        "LISTNAME><LISTWEB>http://cdm2.cemex.com/sites/SSC_HRS/HRA_UK/_vti_bin</LISTWEB><LISTSUBWEB></LISTSUBWEB><ROOTFOLDER" _
'        , _
'        ">/sites/SSC_HRS/HRA_UK/Shared Documents/Employee files/" & EmpFchar & "/" & EmpName & " " & EmpNum & "/Leaver " & Months(LeavingM) & " " & LeavingY & Chr(13) & "" & Chr(10) & "</ROOTFOLDER></LIST>" _
'        )
'        '.CommandType = xlCmdList
'        .connection = _
'        "OLEDB;Provider=Microsoft.Office.List.OLEDB.2.0;Data Source="""";ApplicationName=Excel;Version=12.0.0.0"
'        .RefreshOnFileOpen = False
'        .SavePassword = False
'        .SourceConnectionFile = ""
'        .SourceDataFile = ""
'        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
'        .AlwaysUseConnectionFile = False
'    End With
'    With WB.Connections("Connection")
'        .Name = "Connection"
'        .Description = ""
'    End With
'    WB.Connections("Connection").Refresh

End Sub

Public Sub CommandButton10_Click()
Dim LeaveD As Date
Dim LeaveDL As Date
Dim LeaveDL2 As Date
Dim b As String
Dim wsml As Worksheet, WSCC As Worksheet, WSED As Worksheet
Dim lo As ListObject
Dim dCol As Long
Dim LR As Integer

Set wsml = Application.ThisWorkbook.Worksheets("Managing Leavers")

wsml.Unprotect "bob114"

wsml.Range("b1").CurrentRegion.Sort key1:=wsml.Range("e2"), order1:=xlAscending, Header:=xlYes

If CommandButton10.Caption = "Filtering to this month" Then

'With WSML.OLEObjects("CommandButton8")
'    .Top = WSML.Range("H1").Top
'    .Left = WSML.Range("H1").Left
'    .Width = WSML.Range("H1:J1").Width
'    .Height = WSML.Range("H1").Height
'End With

LR = wsml.Cells(wsml.Rows.Count, 2).End(xlUp).Row


'LeaveDL = DateSerial(Year(Date), Month(Date), 1)
LeaveDL = Month(Date) & "/01/" & Year(Date)
LeaveDL2 = Month(Date) + 1 & "/01/" & Year(Date)

wsml.Range("B1:G" & LR).Autofilter Field:=4, Criteria1:=">=" & LeaveDL, Criteria2:="<" & LeaveDL2
CommandButton10.Caption = "Show All"

Else: GoTo ShowAll

End If
Exit Sub

ShowAll:

wsml.AutoFilterMode = False
CommandButton10.Caption = "Filtering to this month"
wsml.Protect "bob114", DrawingObjects:=False

End Sub

Private Sub CommandButton11_Click()

CommandButton4_Click
CommandButton3_Click

End Sub

Public Sub CommandButton12_Click()

Dim x As Integer, LR As Integer
Dim wb As Workbook


Set SapGuiAuto = GetObject("SAPGUI")
Set Applic = SapGuiAuto.GetScriptingEngine
Set Connection = Applic.Children(0)
Set session = Connection.Children(0)
Set grEmpNumRng = Selection

Set wb = Application.ThisWorkbook

Set WS = wb.Sheets.Add
LR = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row

For Each cell In grEmpNumRng
gsempnumH = cell.Value
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = gsempnumH
    On Error GoTo 0
    session.FindById("wnd[0]").SendVKey 0
    gsEmpName = session.FindById("wnd[0]/usr/subSUBSCR_HEADER:/1PAPAXX/HDR_50060A:0100/txt$_DG01_500A60_DAT_P0001_ENAME").text
    c = c + 2

    session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = 15
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]").SendVKey 20
    Application.Wait (Now + TimeValue("0:00:03"))
    NameandNum
    ScreenShot
Next cell
End Sub

Public Sub ScreenShot()
Dim FSO As FileSystemObject
Dim oSBar As SAPFEWSELib.GuiStatusbar

c = c + 1
Set SapGuiAuto = GetObject("SAPGUI")
Set Applic = SapGuiAuto.GetScriptingEngine
Set Connection = Applic.Children(0)
Set session = Connection.Children(0)


'If session.findById("wnd[0]/sbar").MessageType = "S" Or session.findById("wnd[0]/sbar").MessageType = "E" Then
'Application.Wait Now + 0.00002
'session.findById("wnd[0]").HardCopy "C:\test.jpg", 2
'    With gwksWS.Pictures.Insert("C:\test.jpg", 2)
'    k = k + 1
'        With .ShapeRange
'            .Width = 100
'            .Height = 600
'        End With
'
'    .Left = gwksWS.Cells(c, 1).Left
'    .Top = gwksWS.Cells(c, 1).Top
'    .Placement = 1
'    End With
'    gwksWS.Shapes(k).PictureFormat.CropTop = 650
'    gwksWS.Shapes(k).Left = gwksWS.Cells(c, 1).Left
'    gwksWS.Shapes(k).Top = gwksWS.Cells(c, 1).Top
'    gwksWS.Shapes(k).Placement = 1
'    Kill "C:\test.jpg"
'    c = c + 5
'Else

'    Set oSBar = session.finById("wnd[0]/sbar")
'
'    If oSBar.MessageId = "x" And oSBar.MessageNumber = "123" Then Debug.Print
If Dir("C:\SAPscrFolderxyz", vbDirectory) = "" Then
    MkDir "C:\SAPscrFolderxyz"
End If
    session.FindById("wnd[0]").HardCopy "C:\SAPscrFolderxyz\test.jpg", 1
    'session.FindById("wnd[0]").HardCopy "Documents\test.jpg", 1
    WS.Shapes.AddPicture "C:\SAPscrFolderxyz\test.jpg", msoFalse, True, WS.Cells(c, "j").Left, WS.Cells(c, "j").Top, 750, 600
'    k = k + 1
'        With .ShapeRange
'            .Width = 100
'            .Height = 600
'        End With
'    .Left = gwksWS.Cells(c, 1).Left
'    .Top = gwksWS.Cells(c, 1).Top
'    .Placement = 1
'    End With

    Kill "C:\SAPscrFolderxyz\*.*"
    RmDir "C:\SAPscrFolderxyz"
    c = c + 45
'End If


End Sub

Public Sub NameandNum()

With WS.Range("J" & c & ":K" & c)
    .Merge
    .Value = gsEmpName & " " & gsempnumH
    .Font.Size = 16
    .Font.FontStyle = "Calibri"
End With

End Sub

Private Sub CommandButton14_Click()


Dim SapGuiAuto
Dim SAPApp
Dim SAPCon
Dim session
Dim TodayDay As String, TodayMonth As String, TodayYear As String, TodayDate As String, ChNum As Integer, ChNum2 As Integer, ChNum3 As Integer, ChNum4 As Integer, SupS1 As String, SupS2 As String, SupS3 As String, SupS4 As String
Dim EmpName As String, EmpFName As String, EmpNum As String, EmpAdd1 As String, EmpAdd2 As String, EmpAdd3 As String, EmpAdd4 As String, EmpAdd5 As String, EmpAdd6 As String
Dim LeaveD As Date, sLeaveD As String, LeaveD1 As String, LeaveD2 As String, LeaveD3 As String, LetterD As String, LeaveRe As String, PayDay As String, LetterD1 As String, LetterD2 As String, LetterD3 As String, PayDay1 As String, PayDay2 As String, PayDay3 As String
Dim ManName As String, ManPos As String, CL As Range, OtherC As String
Dim WSED As Worksheet, wsml As Worksheet, ManLongName As String, Textres As String, Textret As String, Textsubres As String, Textsubret As String

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window) on that connection

'IntPos As Integer

Set wsml = ThisWorkbook.Worksheets("Managing Leavers")

EmpNum = ActiveCell.Value

With wsml.Cells
        Set CL = .Find(EmpNum, LookIn:=xlValues)
        If Not CL Is Nothing Then
            wsml.Activate
            CL.Select
        End If
End With

LeaveRe = wsml.Cells(CL.Row, "d")
LeaveD = wsml.Cells(CL.Row, "e")
sLeaveD = Format(LeaveD, "dd.mm.yyyy")
'sLeaveD = Format(LeaveD, "dd.") & Format(LeaveD, "mm.") & Format(LeaveD, "yyyy")
If LeaveRe = "" Then
MsgBox "Please add a leaving reason to the employee!"
Exit Sub
End If

If LeaveRe = "Resignation" Or LeaveRe = "Retirement" Then
Else
MsgBox "The leaving reason should be Resignation or Retirement to create a confirmation letter, please correct!"
Exit Sub
End If

OtherC = Cells(CL.Row, 7).Value

LetterD = InputBox("Please add the date of the letter", "Letter dated")
PayDay = InputBox("Please add the next payday" & " " & "(Leave date: " & sLeaveD & ")", "Next Payday")

If Len(LetterD) > 0 Then
LetterD1 = Split(LetterD, ".")(0)
LetterD2 = Split(LetterD, ".")(1)
LetterD3 = Split(LetterD, ".")(2)
End If

If Left(LetterD1, 1) = "0" Then
    LetterD1 = LetterD1 - "0"
End If

ChNum3 = Len(LetterD1)

    If ChNum3 = 1 Then

        Select Case LetterD1
        Case "1"
        SupS1 = "st"
        Case "2"
        SupS1 = "nd"
        Case "3"
        SupS1 = "rd"
        Case Else
        SupS1 = "th"
        End Select
    ElseIf ChNum3 = 2 Then

        Select Case LetterD1
        Case "21"
        SupS1 = "st"
        Case "22"
        SupS1 = "nd"
        Case "23"
        SupS1 = "rd"
        Case "31"
        SupS1 = "st"
        Case Else
        SupS1 = "th"
        End Select
     End If

        Select Case LetterD2
        Case "01"
        LetterD2 = "January"
        Case "02"
        LetterD2 = "February"
        Case "03"
        LetterD2 = "March"
        Case "04"
        LetterD2 = "April"
        Case "05"
        LetterD2 = "May"
        Case "06"
        LetterD2 = "June"
        Case "07"
        LetterD2 = "July"
        Case "08"
        LetterD2 = "August"
        Case "09"
        LetterD2 = "September"
        Case "10"
        LetterD2 = "October"
        Case "11"
        LetterD2 = "November"
        Case "12"
        LetterD2 = "December"
    End Select


LeaveD1 = Split(sLeaveD, ".")(0)
LeaveD2 = Split(sLeaveD, ".")(1)
LeaveD3 = Split(sLeaveD, ".")(2)
If Left(LeaveD1, 1) = "0" Then
    LeaveD1 = LeaveD1 - "0"
End If
ChNum2 = Len(LeaveD1)
    If ChNum2 = 1 Then

        Select Case LeaveD1
        Case "1"
        SupS4 = "st"
        Case "2"
        SupS4 = "nd"
        Case "3"
        SupS4 = "rd"
        Case Else
        SupS4 = "th"
        End Select
    ElseIf ChNum2 = 2 Then

        Select Case LeaveD1
        Case "21"
        SupS4 = "st"
        Case "22"
        SupS4 = "nd"
        Case "23"
        SupS4 = "rd"
        Case "31"
        SupS4 = "st"
        Case Else
        SupS4 = "th"
        End Select
     End If

        Select Case LeaveD2
        Case "01"
        LeaveD2 = "January"
        Case "02"
        LeaveD2 = "February"
        Case "03"
        LeaveD2 = "March"
        Case "04"
        LeaveD2 = "April"
        Case "05"
        LeaveD2 = "May"
        Case "06"
        LeaveD2 = "June"
        Case "07"
        LeaveD2 = "July"
        Case "08"
        LeaveD2 = "August"
        Case "09"
        LeaveD2 = "September"
        Case "10"
        LeaveD2 = "October"
        Case "11"
        LeaveD2 = "November"
        Case "12"
        LeaveD2 = "December"
    End Select

PayDay1 = Split(PayDay, ".")(0)
PayDay2 = Split(PayDay, ".")(1)
PayDay3 = Split(PayDay, ".")(2)
If Left(PayDay1, 1) = "0" Then
    PayDay1 = PayDay1 - "0"
End If
ChNum4 = Len(PayDay1)
    If ChNum4 = 1 Then

        Select Case PayDay1
        Case "1"
        SupS3 = "st"
        Case "2"
        SupS3 = "nd"
        Case "3"
        SupS3 = "rd"
        Case Else
        SupS3 = "th"
        End Select
    ElseIf ChNum4 = 2 Then

        Select Case PayDay1
        Case "21"
        SupS3 = "st"
        Case "22"
        SupS3 = "nd"
        Case "23"
        SupS3 = "rd"
        Case "31"
        SupS3 = "st"
        Case Else
        SupS3 = "th"
        End Select
     End If

        Select Case PayDay2
        Case "01"
        PayDay2 = "January"
        Case "02"
        PayDay2 = "February"
        Case "03"
        PayDay2 = "March"
        Case "04"
        PayDay2 = "April"
        Case "05"
        PayDay2 = "May"
        Case "06"
        PayDay2 = "June"
        Case "07"
        PayDay2 = "July"
        Case "08"
        PayDay2 = "August"
        Case "09"
        PayDay2 = "September"
        Case "10"
        PayDay2 = "October"
        Case "11"
        PayDay2 = "November"
        Case "12"
        PayDay2 = "December"
    End Select

        On Error Resume Next
        session.FindById("wnd[0]/tbar[0]/okcd").text = "pa20"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
        If Err.Number <> 0 Then
            session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
            session.FindById("wnd[0]").SendVKey 0
            session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
        End If
        session.FindById("wnd[0]").SendVKey 0
        EmpName = session.FindById("wnd[0]/usr/subSUBSCR_HEADER:/1PAPAXX/HDR_50060A:0100/txt$_DG01_500A60_DAT_P0001_ENAME").text
        session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "6"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]").SendVKey 7
        'session.findById("wnd[0]/usr/txtP0006-STRAS").SetFocus
        EmpAdd1 = session.FindById("wnd[0]/usr/txtP0006-STRAS").text

        EmpAdd2 = session.FindById("wnd[0]/usr/txtP0006-LOCAT").text

        EmpAdd3 = session.FindById("wnd[0]/usr/txtP0006-ORT02").text

        EmpAdd4 = session.FindById("wnd[0]/usr/txtP0006-ORT01").text

        EmpAdd5 = session.FindById("wnd[0]/usr/txtP0006-PSTLZ").text

        EmpAdd6 = session.FindById("wnd[0]/usr/txtT005U-BEZEI").text

        session.FindById("wnd[0]").SendVKey 3

        session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU").getAbsoluteRow(1).Selected = True
    session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,1]").SetFocus
    session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,1]").caretPosition = 0
    session.FindById("wnd[0]/tbar[1]/btn[7]").Press
    If session.FindById("wnd[0]/usr/subSUBSCREEN_T582C:ZP000100:0200/txtPA0001-BOSS_ENAME").text = "There is no relationship" Then
    session.FindById("wnd[0]").SendVKey 19
    End If
    ManLongName = session.FindById("wnd[0]/usr/subSUBSCREEN_T582C:ZP000100:0200/txtPA0001-BOSS_ENAME").text

    If Split(ManLongName, " ")(0) = "Mr" Then ManName = Trim(Replace(ManLongName, "Mr", ""))
    If Split(ManLongName, " ")(0) = "Mr." Then ManName = Trim(Replace(ManLongName, "Mr.", ""))

    If Split(ManLongName, " ")(0) = "Mrs" Then ManName = Trim(Replace(ManLongName, "Mrs", ""))
    If Split(ManLongName, " ")(0) = "Mrs." Then ManName = Trim(Replace(ManLongName, "Mrs.", ""))

    If Split(ManLongName, " ")(0) = "Ms" Then ManName = Trim(Replace(ManLongName, "Ms", ""))
    If Split(ManLongName, " ")(0) = "Ms." Then ManName = Trim(Replace(ManLongName, "Ms.", ""))

    If Split(ManLongName, " ")(0) = "Miss" Then ManName = Trim(Replace(ManLongName, "Miss", ""))
    If Split(ManLongName, " ")(0) = "Miss." Then ManName = Trim(Replace(ManLongName, "Miss.", ""))

    ManPos = session.FindById("wnd[0]/usr/subSUBSCREEN_T582C:ZP000100:0200/txtPA0001-BOSS_PLTXT").text

'        ManName = WSED.Cells(cl.Row, 5)
'
'        ManPos = WSED.Cells(cl.Row, 6)

        session.FindById("wnd[0]").SendVKey 3


    'IntPos = InStr(EmpName, " ")
    EmpFName = Split(EmpName, " ")(1)
    TodayDay = Day(Date)
    ChNum = Len(TodayDay)
    If ChNum = 1 Then

        Select Case TodayDay
        Case "1"
        SupS2 = "st"
        Case "2"
        SupS2 = "nd"
        Case "3"
        SupS2 = "rd"
        Case Else
        SupS2 = "th"
        End Select
    ElseIf ChNum = 2 Then

        Select Case TodayDay
        Case "21"
        SupS2 = "st"
        Case "22"
        SupS2 = "nd"
        Case "23"
        SupS2 = "rd"
        Case "31"
        SupS2 = "st"
        Case Else
        SupS2 = "th"
        End Select
     End If


    TodayMonth = Month(Date)

    Select Case TodayMonth
        Case "1"
        TodayMonth = "January"
        Case "2"
        TodayMonth = "February"
        Case "3"
        TodayMonth = "March"
        Case "4"
        TodayMonth = "April"
        Case "5"
        TodayMonth = "May"
        Case "6"
        TodayMonth = "June"
        Case "7"
        TodayMonth = "July"
        Case "8"
        TodayMonth = "August"
        Case "9"
        TodayMonth = "September"
        Case "10"
        TodayMonth = "October"
        Case "11"
        TodayMonth = "November"
        Case "12"
        TodayMonth = "December"
    End Select

    TodayYear = Year(Date)


    TodayDate = TodayDay & " " & TodayMonth & " " & TodayYear



Dim objWord
Dim objSelection
Dim objDoc
Dim ssInitFilePath As String, sDestFolder As String

Set wsUserData = ThisWorkbook.Worksheets("UserData")

  If wsUserData.OLEObjects("Checkbox1").Object.Value = True Then
    sInitFilePath = wsUserData.Cells(6, 2)
    sDestFolder = wsUserData.Cells(7, 2)
  Else
    sInitFilePath = wsUserData.Cells(1, 2)
    sDestFolder = wsUserData.Cells(2, 2)
  End If

  Set objWord = CreateObject("Word.Application")
  Set objDoc = objWord.Documents.Open(sInitFilePath)

   objWord.Visible = True

Set objSelection = objWord.Selection

    objSelection.Font.Name = "Arial"
    'objSelection.TypeText (TodayDay)
    'With objSelection
    '.Font.Superscript = True
    'objSelection.TypeText (SupS2)
    '.Font.Superscript = False
    'End With'
    objSelection.TypeText (TodayDay & " " & TodayMonth & " " & TodayYear)
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    With objSelection
    .Font.Bold = True
    .Font.Underline = True
    objSelection.TypeText text:="PRIVATE & CONFIDENTAL"
    End With
    objSelection.Font.Bold = wdToggle
    objSelection.Font.Underline = wdUnderlineNone
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText (EmpName)
    objSelection.TypeParagraph

    If Len(EmpAdd1) > 0 Then
    objSelection.TypeText (EmpAdd1)
    End If
    If Len(EmpAdd2) > 0 Then
    objSelection.TypeParagraph
    objSelection.TypeText (EmpAdd2)
    End If
    If Len(EmpAdd3) > 0 Then
    objSelection.TypeParagraph
    objSelection.TypeText (EmpAdd3)
    End If
    If Len(EmpAdd4) > 0 Then
    objSelection.TypeParagraph
    objSelection.TypeText (EmpAdd4)
    End If
    If Len(EmpAdd6) > 0 Then
    objSelection.TypeParagraph
    objSelection.TypeText (EmpAdd6)
    End If
    If Len(EmpAdd5) > 0 Then
    objSelection.TypeParagraph
    objSelection.TypeText (EmpAdd5)
    End If

    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText ("Dear " & EmpFName & ",")
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
If LeaveRe = "Resignation" Then
    objSelection.TypeText ("Further to your letter dated " & LetterD1)
    'With objSelection
    ''.Font.Superscript = True
    'objSelection.TypeText (SupS1)
    ''.Font.Superscript = False
    'End With
    objSelection.TypeText (" " & LetterD2 & " " & LetterD3 & " regarding your resignation of employment from CEMEX, I confirm that your final date of employment will be " & LeaveD1)
    'With objSelection
    ''.Font.Superscript = True
    'objSelection.TypeText (SupS4)
    ''.Font.Superscript = False
    'End With
    objSelection.TypeText (" " & LeaveD2 & " " & LeaveD3 & ".")

        ElseIf LeaveRe = "Retirement" Then
    objSelection.TypeText ("Further to your letter dated " & LetterD1)
    'With objSelection
    ''.Font.Superscript = True
    'objSelection.TypeText (SupS1)
    ''.Font.Superscript = False
    'End With
    objSelection.TypeText (" " & LetterD2 & " " & LetterD3 & " regarding your retirement from CEMEX, I confirm that your final date of employment will be " & LeaveD1)
    'With objSelection
    ''objSelection.TypeText (SupS4)
    ''.Font.Superscript = False
    'End With
    objSelection.TypeText (" " & LeaveD2 & " " & LeaveD3 & ".")
End If


    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText ("All outstanding monies will be paid directly into your bank account on " & PayDay1)
    'With objSelection
    ''.Font.Superscript = True
    'objSelection.TypeText (SupS3)
    ''.Font.Superscript = False
    'End With
    objSelection.TypeText (" " & PayDay2 & " " & PayDay3 & " and your P45 will be forwarded to your home address.")
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText ("On behalf of CEMEX, I would like to thank you for all your hard work and wish you well for the future.")
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText ("Should you have any queries, please let me know.")
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText ("Yours sincerely,")
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeParagraph

    With objSelection
    objSelection.TypeText (ManName)
    objSelection.TypeParagraph
    .Font.Bold = True
    objSelection.TypeText (ManPos)
    End With

Textres = "Hello " & Split(ManName, " ")(0) & "," & vbNewLine & vbNewLine & "I have attached a draft letter to confirm " & Split(EmpName, " ")(1) & "'s resignation and leave date. This letter will be sent on your behalf, after you send us back your approval." & vbNewLine & vbNewLine & "Kind regards," & vbNewLine & "Joe"
Textret = "Hello " & Split(ManName, " ")(0) & "," & vbNewLine & vbNewLine & "I have attached a draft letter to confirm " & Split(EmpName, " ")(1) & "'s retirement and leave date. This letter will be sent on your behalf, after you send us back your approval." & vbNewLine & vbNewLine & "Kind regards," & vbNewLine & "Joe"
Textsubres = Split(EmpName, " ")(1) & " " & Split(EmpName, " ")(2) & " - Letter to confirm resignation"
Textsubret = Split(EmpName, " ")(1) & " " & Split(EmpName, " ")(2) & " - Letter to confirm retirement"

If LeaveRe = "Resignation" Then
'    Application.DisplayAlerts

    objDoc.SaveAs (sDestFolder & Split(EmpName, " ")(1) & " " & Split(EmpName, " ")(2) & " Letter to Confirm Resignation.doc")
    UserForm4.TextBox1 = Textres
    UserForm4.TextBox2 = ManName
    UserForm4.TextBox3 = OtherC
    UserForm4.TextBox4 = Textsubres
    UserForm4.Show
        ElseIf LeaveRe = "Retirement" Then
        objDoc.SaveAs (sDestFolder & Split(EmpName, " ")(1) & " " & Split(EmpName, " ")(2) & " Letter to Confirm Retirement.doc")
        UserForm4.TextBox1 = Textret
        UserForm4.TextBox2 = ManName
        UserForm4.TextBox3 = OtherC
        UserForm4.TextBox4 = Textsubret
        UserForm4.Show

End If


End Sub

Private Sub CommandButton15_Click()
Dim x As Integer, LR As Integer
Dim wb As Workbook
Dim it As String
Dim wsml As Worksheet

Set SapGuiAuto = GetObject("SAPGUI")
Set Applic = SapGuiAuto.GetScriptingEngine
Set Connection = Applic.Children(0)
Set session = Connection.Children(0)
Set grEmpNumRng = Selection

Set wb = Application.ThisWorkbook

Set wsml = wb.Worksheets("Managing Leavers")

gsempnumH = ActiveCell.Value
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = gsempnumH
    On Error GoTo 0
    session.FindById("wnd[0]").SendVKey 0
    gsEmpName = session.FindById("wnd[0]/usr/subSUBSCR_HEADER:/1PAPAXX/HDR_50060A:0100/txt$_DG01_500A60_DAT_P0001_ENAME").text

    session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "15"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]").SendVKey 20

End Sub

Public Sub CommandButton16_Click()

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window) on that connection

Set gwksML = ThisWorkbook.Worksheets("Managing Leavers")
Set gwksCL = ThisWorkbook.Worksheets("CDM Leavers")

gsempnum = gwksML.Cells(ActiveCell.Row, "b")

CommandButton18_Click

With gwksML.Cells
        Set grngCl = .Find(gsempnum, LookIn:=xlValues)
        If Not grngCl Is Nothing Then
            gwksML.Activate
            grngCl.Select
        End If
End With


gdLeaveD = gwksML.Cells(grngCl.Row, "e")
'gdLeaveD = Format(gdLeaveD, "dd/mm/yyyy")
'giLyear = Split(gsLeaveD, "/")(2)
'giLmonth = Split(gsLeaveD, "/")(1)
'giLday = Split(gsLeaveD, "/")(0)'
gsLeaveD = Format(gdLeaveD, "dd.mm.yyyy")
'gsLeaveD = Format(gdLeaveD, "dd.") & Format(gdLeaveD, "mm.") & Format(gdLeaveD, "yyyy")
gsLeavingR = gwksML.Cells(grngCl.Row, "d")
If gsLeavingR = "" Then
MsgBox "Leaving reason has not been specified, please update the CDM tracker.", vbOKOnly
Exit Sub
End If
'gdteLeaveDD = Format(DateSerial(giLyear, giLmonth, giLday), "dd/mm/yyyy")
Select Case gsLeavingR
    Case "Redundancy"
    gsLeavingC = "GF"
    Case "Resignation"
    gsLeavingC = "G9"
    Case "Retirement"
    gsLeavingC = "GA"
    Case "Compromise Agreement"
    gsLeavingC = "GK"
    Case "Death"
    gsLeavingC = "05"
    Case "Resigned in Process"
    gsLeavingC = "GN"
    Case "End of Temporary Contract"
    gsLeavingC = "G8"
    Case "Dismissal Gross Misconduct"
    gsLeavingC = "GC"
    Case "Dismissal Capability"
    gsLeavingC = "GD"
    Case "Dismissal Conduct"
    gsLeavingC = "GJ"
    Case "Tupe Transfer"
    gsLeavingC = "GI"
End Select

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = gsempnum
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "0"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 20
gsLeaveDSAP = session.FindById("wnd[0]/usr/tblMP000000TC3000/txtP0000-BEGDA[0,0]").text
giLSyear = Split(gsLeaveDSAP, ".")(2)
giLSmonth = Split(gsLeaveDSAP, ".")(1)
giLSday = Split(gsLeaveDSAP, ".")(0)
'gdteLeaveDSAP = CDate(gsLeaveDSAP.Value)
gdteLeaveDSAP = DateSerial(giLSyear, giLSmonth, giLSday)
If gdteLeaveDSAP = gdLeaveD + 1 And session.FindById("wnd[0]/usr/tblMP000000TC3000/ctxtP0000-MASSN[2,0]").text = "10" Then
MsgBox "The employee is already delimited"
Exit Sub
End If

If session.FindById("wnd[0]/usr/tblMP000000TC3000/ctxtP0000-MASSN[2,0]").text = "10" Then
    If gdteLeaveDSAP = gdLeaveD + 1 Then
    Else
    MsgBox "The employe is already delimited, but with a wrong leave date, please correct!"
    Exit Sub
    End If
End If
session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "2001"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 20
If session.FindById("wnd[0]/sbar").text = "No data stored for Absences in the selected period" Then
Else
gsAbsLeave = session.FindById("wnd[0]/usr/tblMP200000TC3000/txtP2001-ENDDA[1,0]").text
giALday = Split(gsAbsLeave, ".")(0)
giALmonth = Split(gsAbsLeave, ".")(1)
giALyear = Split(gsAbsLeave, ".")(2)
gdteAbsLeave = DateSerial(giALyear, giALmonth, giALday)

If Not session.FindById("wnd[0]/usr/tblMP200000TC3000/txtT554T-ATEXT[3,0]").text = "Maternity Leave" Then
    If gdteAbsLeave > gdLeaveD Then
    MsgBox "There is absence after leave date, please remove it before delimitation!"
    Exit Sub
        ElseIf gdteAbsLeave + 30 > gdLeaveD Then
            If session.FindById("wnd[0]/usr/tblMP200000TC3000/txtT554T-ATEXT[3,0]").text = "Long Term Sick - Unpaid" Or session.FindById("wnd[0]/usr/tblMP200000TC3000/txtT554T-ATEXT[3,0]").text = "Long Term Sick - Paid" Then
                MsgBox "There is Long term sick around leave date - notify Mate!"
            End If
    End If
Else
MsgBox "There is maternity leave after leave date, please check later, the delimitation process is continuing"
End If

End If


'session.findById("wnd[0]").sendVKey 3
session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa40"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = gsempnum
session.FindById("wnd[0]/usr/ctxtRP50G-EINDA").text = gsLeaveD
session.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT").getAbsoluteRow(2).Selected = True
session.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT/txtT529T-MNTXT[0,2]").SetFocus
session.FindById("wnd[0]/usr/tblSAPMP50ATC_MENU_EVENT/txtT529T-MNTXT[0,2]").caretPosition = 0
session.FindById("wnd[0]/tbar[1]/btn[8]").Press

If Not session.FindById("wnd[0]/sbar").text = "" Then
    If session.FindById("wnd[0]/sbar").MessageType = "E" Then
    MsgBox "An SAP error occured, please check"
    End
    End If
End If


If session.ActiveWindow.text = "Copy Actions" Then
CopyAct
End If

If session.ActiveWindow.text = "Copy Organizational Assignment" Then
    session.FindById("wnd[0]").SendVKey 11
    If Not session.FindById("wnd[0]/sbar").text = "" Then
        Do While session.FindById("wnd[0]/sbar").MessageType = "W"
            session.FindById("wnd[0]").SendVKey 0
        Loop
    End If
End If

If session.ActiveWindow.text = "Delimit Recurring Payments/Deductions" Then
RecPay
End If

If session.ActiveWindow.text = "Create Additional Payments" Then
session.FindById("wnd[0]").SendVKey 19
End If

If session.ActiveWindow.text = "Display Absence Quotas" Then
session.FindById("wnd[0]").SendVKey 19
End If

If session.ActiveWindow.text = "Delimit Court Orders GB" Then
CourtO
End If

If session.ActiveWindow.text = "Delimit Pension Funds GB" Then
PensFunds
End If

If session.ActiveWindow.text = "Delimit ESS Settings Remuneration Statement" Then
session.FindById("wnd[0]").SendVKey 19
End If


'CourtO:
'On Error Resume Next
'gsCourtOS = session.findById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").text
'If Not gsCourtOS = "" Then
'
'
'    If session.findById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").text = "CSL" Then
'    session.findById("wnd[0]/usr/tblMP007000TC3000").getAbsoluteRow(0).Selected = True
'    session.findById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").SetFocus
'    session.findById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").caretPosition = 0
'    session.findById("wnd[0]").sendVKey 14
'    If session.findById("wnd[0]/sbar").text = "Enter data for correction period for payroll area GM" Then
'    session.findById("wnd[0]").sendVKey 0
'    End If
'    Else
'    MsgBox "The employee has a court order"
'    session.findById("wnd[0]").sendVKey 19
'    End If
'End If

PA30:
session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa30"
session.FindById("wnd[0]").SendVKey 0

655:
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "655"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[20]").Press
If session.FindById("wnd[0]/sbar").text = "No data stored for ESS Settings Remuneration Statement in the selected period" Then
GoTo 9000
Else
session.FindById("wnd[0]/usr/tblMP065500TC3000/txtP0655-BEGDA[0,0]").SetFocus
session.FindById("wnd[0]/usr/tblMP065500TC3000/txtP0655-BEGDA[0,0]").caretPosition = 0
session.FindById("wnd[0]").SendVKey 6
session.FindById("wnd[0]/usr/chkP0655-ESSONLY").Selected = False
session.FindById("wnd[0]/usr/ctxtP0655-ENDDA").SetFocus
session.FindById("wnd[0]/usr/ctxtP0655-ENDDA").caretPosition = 9
gs655D = session.FindById("wnd[0]/usr/ctxtP0655-BEGDA").text
gdte655D = CDate(Replace(gs655D, ".", "/"))
If gdte655D > DateSerial(Year(gdLeaveD), Month(gdLeaveD), 1 - 1) Then
session.FindById("wnd[0]/usr/ctxtP0655-ENDDA").text = session.FindById("wnd[0]/usr/ctxtP0655-BEGDA").text
Else
session.FindById("wnd[0]/usr/ctxtP0655-ENDDA").text = Format(DateSerial(Year(gdLeaveD), Month(gdLeaveD), 1 - 1), "dd.mm.yyyy")
End If
session.FindById("wnd[0]").SendVKey 0
    If session.FindById("wnd[0]/sbar").text = "Save your entries" Then
        session.FindById("wnd[0]").SendVKey 11
    End If
    If session.FindById("wnd[0]/sbar").text = "No change found" Then
        session.FindById("wnd[0]").SendVKey 3
    End If
session.FindById("wnd[0]").SendVKey 3
End If

9000:
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "9000"
session.FindById("wnd[0]/tbar[1]/btn[20]").Press
    If session.FindById("wnd[0]/sbar").text = "No data stored for List Indicators and characteristics empl in the selected period" Then
    Exit Sub
    End If
session.FindById("wnd[0]/usr/tblMP900000TC3000").getAbsoluteRow(0).Selected = True
session.FindById("wnd[0]/usr/tblMP900000TC3000/txtP9000-BEGDA[0,0]").SetFocus
session.FindById("wnd[0]/usr/tblMP900000TC3000/txtP9000-BEGDA[0,0]").caretPosition = 0
session.FindById("wnd[0]/tbar[1]/btn[6]").Press
session.FindById("wnd[0]/usr/tabsTAB1/tabpTAB1_FC2").Select
session.FindById("wnd[0]/usr/tabsTAB1/tabpTAB1_FC2/ssubTAB1_SCA:MP900000:2002/chkP9000-RH1").Selected = False
session.FindById("wnd[0]/usr/tabsTAB1/tabpTAB1_FC2/ssubTAB1_SCA:MP900000:2002/chkP9000-RH2").Selected = False
session.FindById("wnd[0]/usr/tabsTAB1/tabpTAB1_FC2/ssubTAB1_SCA:MP900000:2002/chkP9000-RH3").Selected = False
session.FindById("wnd[0]/usr/ctxtP9000-ENDDA").text = Format(DateSerial(Year(gdLeaveD), Month(gdLeaveD), Day(gdLeaveD) + 1), "dd.mm.yyyy")
session.FindById("wnd[0]").SendVKey 0
    If session.FindById("wnd[0]/sbar").text = "Save your entries" Then
        session.FindById("wnd[0]").SendVKey 11
    End If
session.FindById("wnd[0]").SendVKey 3
session.FindById("wnd[0]").SendVKey 3

MsgBox "The employee has been delimited successfully"

End Sub
Public Sub CourtO()

    If session.FindById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").text = "CSL" Then
    session.FindById("wnd[0]/usr/tblMP007000TC3000").getAbsoluteRow(0).Selected = True
    session.FindById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").SetFocus
    session.FindById("wnd[0]/usr/tblMP007000TC3000/txtP0070-CRTTY[0,0]").caretPosition = 0
    session.FindById("wnd[0]").SendVKey 14
    Do While session.FindById("wnd[0]/sbar").MessageType = "W"
    session.FindById("wnd[0]").SendVKey 0
    Loop
        If session.FindById("wnd[0]/sbar").MessageType = "E" Then
        MsgBox "Court Order delimitation error!"
        session.FindById("wnd[0]").SendVKey 19
            ElseIf session.FindById("wnd[0]/sbar").MessageType = "S" Then
            MsgBox "Student Loan has been delimited!"
        Else
        MsgBox "The employee has a court order"
        session.FindById("wnd[0]").SendVKey 19
        End If
End If

End Sub
Public Sub PensFunds()
giX = 0
Do While Not session.FindById("wnd[0]/usr/tblMP007100TC3000/ctxtP0071-BEGDA[1," & giX & "]").text = "__________"
session.FindById("wnd[0]/usr/tblMP007100TC3000").getAbsoluteRow(giX).Selected = True
giX = giX + 1
Loop
session.FindById("wnd[0]").SendVKey 14

Do While session.FindById("wnd[0]/sbar").MessageType = "W"
    session.FindById("wnd[0]").SendVKey 0
Loop
End Sub

Public Sub RecPay()
giX = 0
Do While Not session.FindById("wnd[0]/usr/tblMP001400TC3000/ctxtP0014-LGART[0," & giX & "]").text = "____"
session.FindById("wnd[0]/usr/tblMP001400TC3000").getAbsoluteRow(giX).Selected = True
giX = giX + 1
Loop
session.FindById("wnd[0]").SendVKey 13

Do While session.FindById("wnd[0]/sbar").MessageType = "W"
    session.FindById("wnd[0]").SendVKey 0
Loop
End Sub

Public Sub CopyAct()
session.FindById("wnd[0]/usr/ctxtP0000-MASSG").text = gsLeavingC
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 11
Do While session.FindById("wnd[0]/sbar").MessageType = "W"
    session.FindById("wnd[0]").SendVKey 0
Loop
If session.ActiveWindow.Name = "wnd[1]" Then
    If session.ActiveWindow.text = "Create Vacancy" Then
    session.FindById("wnd[1]").SendVKey 12
    End If
End If
End Sub



Public Sub OptionButt()
Dim gsDismq As String
UserForm8.Show
    If UserForm8.OptionButton1.Value = True Then
    gsDismq = MsgBox("Reason is 'Dismissal gross misconduct', are you sure?", vbYesNo)
        If gsDismq = vbYes Then
        gsLeavingC = "GC"
        End If
    End If
    If UserForm8.OptionButton2.Value = True Then
    gsDismq = MsgBox("Reason is 'Dismissal conduct', are you sure?", vbYesNo)
        If gsDismq = vbYes Then
        gsLeavingC = "GJ"
        End If
    End If
    If UserForm8.OptionButton3.Value = True Then
    gsDismq = MsgBox("Reason is 'Dismissal capability', are you sure?", vbYesNo)
        If gsDismq = vbYes Then
        gsLeavingC = "GD"
        End If
    End If

End Sub


Private Sub CommandButton17_Click()

Dim EmpNum As String, LeavingR As String, EmpName As String
Dim CL As Range, FolderP As String
Dim WSED As Worksheet, WSCHP As Worksheet, wsml As Worksheet

Set wsml = ThisWorkbook.Worksheets("Managing Leavers")

EmpNum = ActiveCell.Value

With wsml
Set CL = .Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
If Not CL Is Nothing Then CL.Select
End With

EmpName = wsml.Cells(CL.Row, "c")


UserForm9.TextBox1 = "GB HR"
UserForm9.TextBox2 = "Position closing - " & EmpName
UserForm9.TextBox3 = "Hi Team," & vbNewLine & vbNewLine & "Could you please advise if we can close the position of " & EmpName & " " & EmpNum & " as it was indicated in the HRLF that he/she will not be replaced." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & "Joe"

UserForm9.Show

wsml.Activate

End Sub


Public Sub CommandButton18_Click()

Dim wsml As Worksheet, WSCL As Worksheet, WSED As Worksheet, EmpCol As Range, DateCol As Range, ReasonCol As Range, NameCol As Range, Copy As Range
Dim LRCL As Integer, LRML As Integer, Autofil As Integer
Dim FilBut As OLEObject, RNG As Range, firstColor As Long, secondColor As Long

Set wsml = Application.ThisWorkbook.Worksheets("Managing Leavers")
Set WSCL = Application.ThisWorkbook.Worksheets("CDM Leavers")
Set FilBut = wsml.OLEObjects("CommandButton10")

wsml.Unprotect "bob114"

Application.ScreenUpdating = False

If wsml.AutoFilterMode = True Then
Autofil = 1
wsml.AutoFilterMode = False
FilBut.Object.Caption = "Filtering to this month"
End If
ThisWorkbook.RefreshAll

LRML = wsml.Cells(Rows.Count, 2).End(xlUp).Row
LRCL = WSCL.Cells(Rows.Count, 7).End(xlUp).Row
Set EmpCol = WSCL.Columns.Range("A1:X1").Find("Employee Number")
Set DateCol = WSCL.Columns.Range("A1:X1").Find("Leave date")
Set ReasonCol = WSCL.Columns.Range("A1:X1").Find("Reason For Leaving")
Set NameCol = WSCL.Columns.Range("A1:X1").Find("Employee Name")

wsml.Range(wsml.Cells(2, "b"), wsml.Cells(LRML, "e")).ClearContents

WSCL.Activate

'WSCL.Range(WSCL.Cells(2, DateCol.Column), WSCL.Cells(LRCL, DateCol.Column)).NumberFormat = "dd.mm.yyyy"

WSCL.Range(WSCL.Cells(2, EmpCol.Column), WSCL.Cells(LRCL, EmpCol.Column)).Copy
wsml.Cells(2, 2).PasteSpecial xlPasteAll

WSCL.Range(WSCL.Cells(2, NameCol.Column), WSCL.Cells(LRCL, NameCol.Column)).Copy
wsml.Cells(2, 3).PasteSpecial xlPasteAll

WSCL.Range(WSCL.Cells(2, ReasonCol.Column), WSCL.Cells(LRCL, ReasonCol.Column)).Copy
wsml.Cells(2, 4).PasteSpecial xlPasteAll

WSCL.Range(WSCL.Cells(2, DateCol.Column), WSCL.Cells(LRCL, DateCol.Column)).Copy
wsml.Cells(2, 5).PasteSpecial xlPasteAll

LRML = wsml.Cells(Rows.Count, 2).End(xlUp).Row
wsml.Activate
wsml.Range("B2:E" & LRML).Sort key1:=wsml.Range("E1"), order1:=xlAscending
LRML = wsml.Cells(Rows.Count, 2).End(xlUp).Row
wsml.Range("B2:E" & LRML).Interior.Color = xlNone
If Autofil = 1 Then
Call CommandButton10_Click
End If

Set RNG = wsml.Range("B2:E" & LRML)

Call GreenBarMe(RNG, firstColor, secondColor)
wsml.Cells(2, 2).Select
Application.ScreenUpdating = True

wsml.Protect "bob114", DrawingObjects:=False
End Sub

Public Sub GreenBarMe(RNG As Range, firstColor As Long, secondColor As Long)

    firstColor = vbWhite
    secondColor = RGB(204, 236, 255)

    RNG.Interior.ColorIndex = xlNone
    RNG.FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0"
    RNG.FormatConditions(1).Interior.Color = firstColor
    RNG.FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)<>0"
    RNG.FormatConditions(2).Interior.Color = secondColor
End Sub

Private Sub CommandButton19_Click()
Dim x As Integer, LR As Integer
Dim wb As Workbook
Dim it As String
Dim wsml As Worksheet, gsempnumH As String

Set SapGuiAuto = GetObject("SAPGUI")
Set Applic = SapGuiAuto.GetScriptingEngine
Set Connection = Applic.Children(0)
Set session = Connection.Children(0)
Set grEmpNumRng = Selection

Set wb = Application.ThisWorkbook


Set wsml = wb.Worksheets("Managing Leavers")

gsempnumH = ActiveCell.Value
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = gsempnumH
    session.FindById("wnd[0]").SendVKey 0
End Sub

Private Sub CommandButton2_Click()


Dim x As Integer, y As Integer, z As Integer
Dim EmpName As String, EmpEmail As String, EmpHphone As String, EmpMphone As String, EmpNum As String
Dim NiNum As String
Dim EmpAdd1 As String, EmpAdd2 As String, EmpAdd3 As String, EmpAdd4 As String, EmpAdd5 As String, EmpAdd6 As String
Dim ManName As String, StartD As String, CompCar As String, CarAll As String, LeavingR As String, LeaveDs As String, LeaveD2 As Date
Dim CoCode As String, CostC As String, CL As Range
Dim objSelection, sDestFolder As String, sInitFilePath As String
Dim HRSelection
Dim objDoc
Dim HRDoc
Dim objWord
Dim Docname As String
Dim SapGuiAuto
Dim SAPApp
Dim SAPCon
Dim session

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window) on that connection

Worksheets("Managing Leavers").Activate
EmpNum = ActiveCell.Value
LeaveD2 = Cells(ActiveCell.Row, "e").Value
Set wsUserData = ThisWorkbook.Worksheets("UserData")

If wsUserData.OLEObjects("Checkbox1").Object.Value = True Then
sDestFolder = wsUserData.Cells(10, 2)
sInitFilePath = wsUserData.Cells(9, 2)
Else
sDestFolder = wsUserData.Cells(5, 2)
sInitFilePath = wsUserData.Cells(4, 2)
End If

On Error Resume Next

session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
session.FindById("wnd[0]").SendVKey 0

session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum

If Err.Number <> 0 Then
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
End If

session.FindById("wnd[0]").SendVKey 0

EmpName = session.FindById("wnd[0]/usr/subSUBSCR_HEADER:/1PAPAXX/HDR_50060A:0100/txt$_DG01_500A60_DAT_P0001_ENAME").text


Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open(sInitFilePath)
objWord.Visible = True
Set objSelection = objWord.Selection

With objSelection.Find
    .MatchWholeWord = True
    .MatchCase = False
    .text = "Driver Name"
    .Execute
End With
    objSelection.endof
    objSelection.MoveRight Count:=1
    objSelection.text = EmpName
    objSelection.endof
    objSelection.MoveDown Count:=1
    objSelection.text = EmpNum

session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "41"
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").caretPosition = 2
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[7]").Press
StartD = session.FindById("wnd[0]/usr/ctxtP0041-DAT01").text
session.FindById("wnd[0]").SendVKey 3

 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU").getAbsoluteRow(2).Selected = True
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,2]").SetFocus
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,2]").caretPosition = 0
 session.FindById("wnd[0]").SendVKey 7
 NiNum = session.FindById("wnd[0]/usr/txtQ0002-PERID").text
 session.FindById("wnd[0]").SendVKey 3

         objSelection.MoveDown Count:=1
         objSelection.text = NiNum

 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "105"
 session.FindById("wnd[0]").SendVKey 0
 session.FindById("wnd[0]").SendVKey 20
 If session.FindById("wnd[0]/sbar").text = "No data stored for Communication in the selected period" Then GoTo NextTask

 x = 0

 Do While Not session.FindById("wnd[0]/usr/tblMP010500TC3000/txtP0105-ENDDA[1," & x & "]").text = ""
     If session.FindById("wnd[0]/usr/tblMP010500TC3000/txtP0105-ENDDA[1," & x & "]").text = "__________" Then
     session.FindById("wnd[0]").SendVKey 3
     GoTo NextTask

     End If
     If session.FindById("wnd[0]/usr/tblMP010500TC3000/txtP0105-ENDDA[1," & x & "]").text = "31.12.9999" Then
     If session.FindById("wnd[0]/usr/tblMP010500TC3000/txtT591S-STEXT[3," & x & "]").text = "E-mail" Then
     EmpEmail = session.FindById("wnd[0]/usr/tblMP010500TC3000/txtP0105-USRID_LONG[4," & x & "]").text
         ElseIf session.FindById("wnd[0]/usr/tblMP010500TC3000/txtT591S-STEXT[3," & x & "]").text = "House Telephone" Then
         EmpHphone = session.FindById("wnd[0]/usr/tblMP010500TC3000/txtP0105-USRID_LONG[4," & x & "]").text
             ElseIf session.FindById("wnd[0]/usr/tblMP010500TC3000/txtT591S-STEXT[3," & x & "]").text = "Mobile" Or session.FindById("wnd[0]/usr/tblMP010500TC3000/txtT591S-STEXT[3," & x & "]").text = "Personal Mobile" Then
             EmpMphone = session.FindById("wnd[0]/usr/tblMP010500TC3000/txtP0105-USRID_LONG[4," & x & "]").text
     End If
     End If

 x = x + 1
 If x = 22 Then
 session.FindById("wnd[0]/usr/tblMP010500TC3000").verticalScrollbar.Position = 20
 x = 0
 End If

 Loop

NextTask:

If EmpEmail <> "" Then
objSelection.MoveDown Count:=1
objSelection.text = EmpEmail
Else
objSelection.MoveDown Count:=1
objSelection.text = "-"
End If

If EmpHphone <> "" Then
objSelection.MoveDown Count:=1
objSelection.text = EmpHphone
Else
objSelection.MoveDown Count:=1
objSelection.text = "-"
End If

If EmpMphone <> "" Then
objSelection.MoveDown Count:=1
objSelection.text = EmpMphone
Else
objSelection.MoveDown Count:=1
objSelection.text = "-"
End If

 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,3]").SetFocus
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,3]").caretPosition = 0
 session.FindById("wnd[0]").SendVKey 7
 EmpAdd1 = session.FindById("wnd[0]/usr/txtP0006-STRAS").text

 EmpAdd2 = session.FindById("wnd[0]/usr/txtP0006-LOCAT").text

 EmpAdd3 = session.FindById("wnd[0]/usr/txtP0006-ORT02").text

 EmpAdd4 = session.FindById("wnd[0]/usr/txtP0006-ORT01").text

 EmpAdd5 = session.FindById("wnd[0]/usr/txtP0006-PSTLZ").text

 EmpAdd6 = session.FindById("wnd[0]/usr/txtT005U-BEZEI").text


 objSelection.MoveDown Count:=1

 If EmpAdd1 <> "" Then
 objSelection.text = EmpAdd1
 objSelection.endof
 objSelection.TypeParagraph
 End If

 If EmpAdd2 <> "" Then
 objSelection.text = EmpAdd2
 objSelection.endof
 objSelection.TypeParagraph
 End If

 If EmpAdd3 <> "" Then
 objSelection.text = EmpAdd3
 objSelection.endof
 objSelection.TypeParagraph
 End If

 If EmpAdd4 <> "" Then
 objSelection.text = EmpAdd4
 objSelection.endof
 objSelection.TypeParagraph
 End If

 If EmpAdd5 <> "" Then
 objSelection.text = EmpAdd5
 objSelection.endof
 objSelection.TypeParagraph
 End If

 If EmpAdd6 <> "" Then
 objSelection.text = EmpAdd6
 objSelection.endof
 objSelection.TypeParagraph
 End If

 session.FindById("wnd[0]").SendVKey 3
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU").getAbsoluteRow(1).Selected = True
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,1]").SetFocus
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,1]").caretPosition = 0
 session.FindById("wnd[0]").SendVKey 7

 If session.FindById("wnd[0]/usr/subSUBSCREEN_T582C:ZP000100:0200/txtPA0001-BOSS_ENAME").text = "There is no relationship" Then
 session.FindById("wnd[0]").SendVKey 19
 End If

 ManName = session.FindById("wnd[0]/usr/subSUBSCREEN_T582C:ZP000100:0200/txtPA0001-BOSS_ENAME").text
 CoCode = session.FindById("wnd[0]/usr/ctxtP0001-BUKRS").text
 CostC = session.FindById("wnd[0]/usr/ctxtP0001-KOSTL").text

 objSelection.MoveDown Count:=1
 objSelection.text = CoCode
 objSelection.MoveDown Count:=1
 objSelection.text = CostC

 session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
 session.FindById("wnd[0]").SendVKey 0
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "14"
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").SetFocus
 session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").caretPosition = 2
 session.FindById("wnd[0]").SendVKey 0
 session.FindById("wnd[0]").SendVKey 20

 If session.FindById("wnd[0]/sbar").text = "No data stored for Recurring Payments/Deductions in the selected period" Then GoTo NextTask1

 y = 0

 Do While Not session.FindById("wnd[0]/usr/tblMP001400TC3000/txtP0014-ENDDA[3," & y & "]").text = "__________"

 If Not CompCar = "" Then GoTo NextTask1

 If session.FindById("wnd[0]/usr/tblMP001400TC3000/txtP0014-ENDDA[3," & y & "]").text = "__________" Then
 session.FindById("wnd[0]").SendVKey 3

 GoTo NextTask1

 End If

 If session.FindById("wnd[0]/usr/tblMP001400TC3000/txtP0014-ENDDA[3," & y & "]").text = Replace(LeaveD2, "/", ".") Or session.FindById("wnd[0]/usr/tblMP001400TC3000/txtP0014-ENDDA[3," & y & "]").text = "31.12.9999" Or session.FindById("wnd[0]/usr/tblMP001400TC3000/txtP0014-ENDDA[3," & y & "]").text > Replace(LeaveD2, "/", ".") Then
     If session.FindById("wnd[0]/usr/tblMP001400TC3000/txtT512T-LGTXT[1," & y & "]").text = "Eligible for Company Car" Then
     session.FindById("wnd[0]/usr/tblMP001400TC3000").getAbsoluteRow(y).Selected = True
     session.FindById("wnd[0]").SendVKey 2

     CompCar = session.FindById("wnd[0]/usr/txtP0014-ZUORD").text

     session.FindById("wnd[0]").SendVKey 3

         ElseIf session.FindById("wnd[0]/usr/tblMP001400TC3000/txtT512T-LGTXT[1," & y & "]").text = "Eligible for Company Van" Then
             CompCar = "VAN"

                 ElseIf session.FindById("wnd[0]/usr/tblMP001400TC3000/txtT512T-LGTXT[1," & y & "]").text = "Car Allowance" Then
                 CarAll = session.FindById("wnd[0]/usr/tblMP001400TC3000/txtQ0014-BETRG[5," & y & "]").text
     End If
 End If

 If y = 22 Then
 session.FindById("wnd[0]/usr/tblMP010500TC3000").verticalScrollbar.Position = 20
 y = 0
 End If

 y = y + 1
 Loop

NextTask1:

With Worksheets("Managing Leavers").Cells
        Set CL = .Find(EmpNum, LookIn:=xlValues)
        If Not CL Is Nothing Then
            Worksheets("Closing Cases").Activate
            CL.Select
        End If
End With

'StartD = Worksheets("Employee Data").Cells(cl.Row, "k").Value
LeavingR = Worksheets("Managing Leavers").Cells(CL.Row, "d").Value

session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
session.FindById("wnd[0]").SendVKey 0

objSelection.MoveDown Count:=1
objSelection.text = CompCar

If CarAll <> "" Then
objSelection.MoveDown Count:=1
objSelection.text = CarAll
Else
objSelection.MoveDown Count:=1
End If

objSelection.MoveDown Count:=1
objSelection.text = ManName

objSelection.MoveDown Count:=1
objSelection.text = StartD

objSelection.MoveDown Count:=1
objSelection.text = Replace(LeaveD2, "/", ".")

objSelection.MoveDown Count:=1
objSelection.text = LeavingR

'End If

objDoc.SaveAs (sDestFolder)

'End If

Worksheets("Managing Leavers").Activate

End Sub

Public Sub CommandButton13_Click()

Dim x As Integer, LR As Integer
Dim wb As Workbook
Dim it As String
Dim wsml As Worksheet

Set SapGuiAuto = GetObject("SAPGUI")
Set Applic = SapGuiAuto.GetScriptingEngine
Set Connection = Applic.Children(0)
Set session = Connection.Children(0)
Set grEmpNumRng = Selection

Set wb = Application.ThisWorkbook

Set WS = wb.Sheets.Add
LR = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
Set wsml = wb.Worksheets("Managing Leavers")
it = wsml.Shapes("TextBox1").OLEFormat.Object.Object.Value
For Each cell In grEmpNumRng
gsempnumH = cell.Value
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = gsempnumH
    On Error GoTo 0
    session.FindById("wnd[0]").SendVKey 0
    gsEmpName = session.FindById("wnd[0]/usr/subSUBSCR_HEADER:/1PAPAXX/HDR_50060A:0100/txt$_DG01_500A60_DAT_P0001_ENAME").text
    c = c + 2

    session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = it
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]").SendVKey 20
    Application.Wait (Now + TimeValue("0:00:03"))
    NameandNum
    ScreenShot
Next cell


End Sub

Private Sub CommandButton20_Click()
bConnectToSap
End Sub

Private Sub CommandButton3_Click()

Dim Link As String
Dim wsml As Worksheet
Dim CL As Range
Dim EmpNum As String, EmpName As String, EmpFirstName As String, EmpFchar As String, EmpLastName As String, EmpLastName2 As String
Dim NameRange As Range

Set wsml = Application.ThisWorkbook.Worksheets("Managing Leavers")

EmpNum = ActiveCell.Value

Set NameRange = wsml.Range("A1:K1").Find("Employee Name")

With wsml
    Set CL = .Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
End With
    EmpName = wsml.Cells(CL.Row, "c")
    EmpFirstName = Split(EmpName, " ")(0)
    EmpFchar = Left(EmpFirstName, 1)
    EmpLastName = Split(EmpName, " ")(1)
    On Error Resume Next
    EmpLastName2 = Split(EmpName, " ")(2)

If Not EmpLastName2 = "" Then
Link = "http://cdm2.cemex.com/sites/SSC_HRS/HRA_UK/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSSC_HRS%2FHRA_UK%2FShared%20Documents%2FEmployee%20files%2F" & EmpFchar & "%2F" & EmpFirstName & "%20" & EmpLastName & "%20" & EmpLastName2 & "%20" & EmpNum & "&FolderCTID=0x012000DDB6B20617D10F45B07F23329EF84FE3&View=%7B42EEE4CF-14EF-4B58-B95D-61924126CBE2%7D"
Else
Link = "http://cdm2.cemex.com/sites/SSC_HRS/HRA_UK/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSSC_HRS%2FHRA_UK%2FShared%20Documents%2FEmployee%20files%2F" & EmpFchar & "%2F" & EmpFirstName & "%20" & EmpLastName & "%20" & EmpNum & "&FolderCTID=0x012000DDB6B20617D10F45B07F23329EF84FE3&View=%7B42EEE4CF-14EF-4B58-B95D-61924126CBE2%7D"
End If
ActiveWorkbook.FollowHyperlink Address:=Link

End Sub

Private Sub CommandButton4_Click()

Dim Link As String, ID As String, EmpNum As String
Dim WSED As Worksheet, WSCL As Worksheet, wsml As Worksheet
Dim CL As Range
Dim sfindVal As String

EmpNum = ActiveCell.Value
Set wsml = Application.ThisWorkbook.Worksheets("Managing Leavers")
Set WSCL = Application.ThisWorkbook.Worksheets("CDM Leavers")

With WSCL
    Set CL = .Range("G:G").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
End With
On Error GoTo Errormess
ID = WSCL.Cells(CL.Row, "v").Value
Link = "http://cdm2.cemex.com/sites/SSC_HRS/HRA_UK/Lists/Separation%20Tracker/DispForm.aspx?ID=" & ID & "&Source=http%3A%2F%2Fcdm2%2Ecemex%2Ecom%2Fsites%2FSSC_HRS%2FHRA_UK%2FLists%2FSeparation%2520Tracker%2FNYLT%2520GYEK%2Easpx&ContentTypeId=0x01008492981C8E844740AAC610185A76F3C3"

ActiveWorkbook.FollowHyperlink Address:=Link

Exit Sub

Errormess:
MsgBox ("The employee number has not been found in the leaver tracker list!")
End Sub

Private Sub CommandButton5_Click()
Dim WSED As Worksheet, wsml As Worksheet, WSCC As Worksheet, WSCL As Worksheet
Dim sfindVal As String
Dim CL As Range
Dim cl2 As Range
Dim x
Dim LNF As String, HRLF As String, EIF As String, TermLett As String, LeavingR As String, EmpNum As String
Dim EmpName As String, EmpFirstName As String, ManName As String, ManFirstName As String
Dim ChaseText As String, ChaseText2 As String

EmpNum = ActiveCell.Value

Call CommandButton18_Click

Set wsml = ThisWorkbook.Worksheets("Managing Leavers")
Set WSCL = ThisWorkbook.Worksheets("CDM Leavers")
WSCL.Activate

With WSCL
Set CL = .Range("G:G").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
End With

wsml.Activate

With wsml
Set cl2 = .Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
End With

UserForm1.Label1 = wsml.Cells(cl2.Row, 3)

LNF = WSCL.Cells(CL.Row, "j").Value
If LNF = "In progress" Then
LNF = ""
End If
HRLF = WSCL.Cells(CL.Row, "k").Value
If HRLF = "In progress" Then
HRLF = ""
End If
EIF = WSCL.Cells(CL.Row, "l").Value
If EIF = "In progress" Then
EIF = ""
End If
TermLett = WSCL.Cells(CL.Row, "i")
If TermLett = "In progress" Then
TermLett = ""
End If
LeavingR = WSCL.Cells(CL.Row, "h")
EmpNum = WSCL.Cells(CL.Row, "g")

EmpName = WSCL.Cells(CL.Row, "c").Value
EmpFirstName = Split(EmpName)(0)
ManName = WSCL.Cells(CL.Row, "f")
If Not ManName = "" Then
ManFirstName = Split(ManName)(0)
Else
ManFirstName = ""
End If

If LeavingR = "Resignation" Then

If Len(LNF) > 0 And Len(HRLF) > 0 And Len(EIF) > 0 Then
MsgBox ("Nothing to chase")
GoTo EndSub
    ElseIf LNF = "" And HRLF = "" And EIF = "" Then
    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & "Could you please also send the Exit Interview Form, if the interview was conducted?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
        ElseIf LNF = "" And HRLF = "" And Len(EIF) > 0 Then
        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
            ElseIf LNF = "" And Len(HRLF) > 0 And Len(EIF) > 0 Then
            ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                ElseIf Len(LNF) > 0 And Len(HRLF) > 0 And EIF = "" Then
                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "Could you please send me the Exit Interview Form about " & EmpFirstName & ", if the interview was conducted?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                    ElseIf LNF = "" And Len(HRLF) > 0 And EIF = "" Then
                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & "Could you please also send the Exit Interview Form, if the interview was conducted?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                        ElseIf Len(LNF) > 0 And HRLF = "" And Len(EIF) > 0 Then
                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                            ElseIf Len(LNF) > 0 And HRLF = "" And EIF = "" Then
                            ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form about " & EmpFirstName & "?" & vbNewLine & "Could you please also send me the Exit Interview Form, if the interview was conducted?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
End If
End If

If LeavingR = "End of Temporary Contract" Or LeavingR = "Retirement" Or LeavingR = "Tupe Transfer" Or LeavingR = "Resigned in Process" Then

If Len(LNF) > 0 And Len(HRLF) Then
MsgBox ("Nothing to chase")
Exit Sub
    ElseIf LNF = "" And HRLF = "" Then
    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
            ElseIf LNF = "" And Len(HRLF) > 0 Then
            ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                        ElseIf Len(LNF) > 0 And HRLF = "" Then
                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"

End If
End If


If LeavingR = "Redundancy" Then
If Len(LNF) > 0 And Len(HRLF) > 0 And Len(TermLett) > 0 Then
MsgBox ("Nothing to chase")
GoTo EndSub
                                ElseIf LNF = "" And HRLF = "" And TermLett = "" Then
                                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form, Leaver Notification Form and termination letter for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                    ElseIf LNF = "" And HRLF = "" And Len(TermLett) > 0 Then
                                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                        ElseIf LNF = "" And Len(HRLF) > 0 And Len(TermLett) > 0 Then
                                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                            ElseIf Len(LNF) > 0 And Len(HRLF) > 0 And TermLett = "" Then
                                            ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "Could you please send me the termination letter for " & EmpFirstName & " ?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                ElseIf LNF = "" And Len(HRLF) > 0 And TermLett = "" Then
                                                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form and the termination letter for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                    ElseIf Len(LNF) > 0 And HRLF = "" And Len(TermLett) > 0 Then
                                                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                        ElseIf Len(LNF) > 0 And HRLF = "" And TermLett = "" Then
                                                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and termination letter for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"


End If
End If

If LeavingR = "Dismissal" Or LeavingR = "Dismissal Conduct" Or LeavingR = "Dismissal Gross Misconduct" Or LeavingR = "Dismissal Capability" Then

If Len(LNF) > 0 And Len(HRLF) > 0 And Len(TermLett) > 0 Then
MsgBox ("Nothing to chase")
GoTo EndSub
    ElseIf LNF = "" And HRLF = "" And TermLett = "" Then
                                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form, Leaver Notification Form and dismissal letter for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                    ElseIf LNF = "" And HRLF = "" And Len(TermLett) > 0 Then
                                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                        ElseIf LNF = "" And Len(HRLF) > 0 And Len(TermLett) > 0 Then
                                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                            ElseIf Len(LNF) > 0 And Len(HRLF) > 0 And TermLett = "" Then
                                            ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "Could you please send me the dismissal letter for " & EmpFirstName & " ?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                ElseIf LNF = "" And Len(HRLF) > 0 And TermLett = "" Then
                                                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form and the dismissal letter for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                    ElseIf Len(LNF) > 0 And HRLF = "" And Len(TermLett) > 0 Then
                                                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                        ElseIf Len(LNF) > 0 And HRLF = "" And TermLett = "" Then
                                                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and dismissal letter for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
End If
End If


If LeavingR = "Compromise Agreement" Then

If Len(LNF) > 0 And Len(HRLF) > 0 And Len(TermLett) > 0 Then
MsgBox ("Nothing to chase")
GoTo EndSub
    ElseIf LNF = "" And HRLF = "" And TermLett = "" Then
                                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form, Leaver Notification Form and settlement agreement for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                    ElseIf LNF = "" And HRLF = "" And Len(TermLett) > 0 Then
                                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                        ElseIf LNF = "" And Len(HRLF) > 0 And Len(TermLett) > 0 Then
                                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                            ElseIf Len(LNF) > 0 And Len(HRLF) > 0 And TermLett = "" Then
                                            ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "Could you please send me the settlement agreement for " & EmpFirstName & " ?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                ElseIf LNF = "" And Len(HRLF) > 0 And TermLett = "" Then
                                                ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed Leaver Notification Form and the settlement agreement for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                    ElseIf Len(LNF) > 0 And HRLF = "" And Len(TermLett) > 0 Then
                                                    ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form about " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                                                        ElseIf Len(LNF) > 0 And HRLF = "" And TermLett = "" Then
                                                        ChaseText = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "May I ask you to send us the completed HR Leaver Form and the settlement agreement for " & EmpFirstName & "?" & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
End If
End If

UserForm1.TextBox1 = ChaseText
UserForm1.TextBox2 = ManName
'UserForm1.TextBox3 = WSED.Cells(cl.Row, 7).Value
UserForm1.TextBox4 = wsml.Cells(cl2.Row, 3).Value & " - Leaver Forms"
UserForm1.Show

EndSub:

End Sub

Private Sub CommandButton6_Click()

Dim wb As Workbook
Dim wsml As Worksheet, WSLF As Worksheet, WSCL As Worksheet
Dim arrTxt, vTxt, arrLabels, LTxt
Dim ZZ As Integer
Dim EmpName As String, EmpNum As String, EmpFchar As String
Dim objSelection
Dim HRSelection
Dim objDoc
Dim HRDoc
Dim objWord, MadeLeaver As String, EmpNumRange As Range
Dim Docname As String, LeaveD As String, LeavingM As String, LeavingY As String, Months(1 To 12) As String
Dim CL As Range, cl2 As Range, pos As Integer, pos2 As Integer, cl3 As Range, pos3 As Integer, pos4 As Integer
Dim sfindVal As String
Dim Texttofind As String
Dim arrLabvar(0 To 11) As String
Dim z As Integer
Dim FileName As String
Dim filepick As Office.FileDialog
Dim filepick2 As String
Dim SelectedFile As String
Dim clipboard As MSForms.DataObject
Dim wdStory
Dim text As String
Dim sDestFolder As String, sInitFilePath As String, sPathHRLF As String

Set clipboard = New MSForms.DataObject

Set wb = Application.ThisWorkbook
sfindVal = ActiveCell.Value
EmpNum = sfindVal
Set wsml = ThisWorkbook.Worksheets("Managing Leavers")
Set WSLF = ThisWorkbook.Worksheets("Leaver Folder")
Set WSCL = ThisWorkbook.Worksheets("CDM Leavers")
Set wsUserData = ThisWorkbook.Worksheets("UserData")

If wsUserData.OLEObjects("Checkbox1").Object.Value = True Then
  sInitFilePath = wsUserData.Cells(6, 2)
  sDestFolder = wsUserData.Cells(8, 2)
  sPathHRLF = wsUserData.Cells(7, 2)
Else
  sInitFilePath = wsUserData.Cells(1, 2)
  sDestFolder = wsUserData.Cells(3, 2)
  sPathHRLF = wsUserData.Cells(2, 2)
End If

wsml.Activate

With wsml
Set CL = .Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
If Not CL Is Nothing Then CL.Select
End With

EmpNum = sfindVal

EmpName = wsml.Cells(CL.Row, "c")
EmpFchar = Left(EmpName, 1)
LeaveD = wsml.Cells(CL.Row, "e")
LeavingM = Split(LeaveD, "/")(1)
LeavingY = Split(LeaveD, "/")(2)

'EmpNumRange = WSCL.Range("A1:Z1").Find("Employee Number")
Set cl3 = WSCL.Range("G:G").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
MadeLeaver = WSCL.Cells(cl3.Row, "w")

Months(1) = "January"
Months(2) = "February"
Months(3) = "March"
Months(4) = "April"
Months(5) = "May"
Months(6) = "June"
Months(7) = "July"
Months(8) = "August"
Months(9) = "September"
Months(10) = "October"
Months(11) = "November"
Months(12) = "December"

    With wb.Connections("Connection").OLEDBConnection
        .BackgroundQuery = False
        .CommandText = Array( _
        "<LIST><VIEWGUID>{42EEE4CF-14EF-4B58-B95D-61924126CBE2}</VIEWGUID><LISTNAME>{B8910762-E3D1-4AA2-A04E-7D516CED1381}</" _
        , _
        "LISTNAME><LISTWEB>http://cdm2.cemex.com/sites/SSC_HRS/HRA_UK/_vti_bin</LISTWEB><LISTSUBWEB></LISTSUBWEB><ROOTFOLDER" _
        , _
        ">/sites/SSC_HRS/HRA_UK/Shared Documents/Employee files/" & EmpFchar & "/" & EmpName & " " & EmpNum & "/Leaver " & Months(LeavingM) & " " & LeavingY & Chr(13) & "" & Chr(10) & "</ROOTFOLDER></LIST>" _
        )
        '.CommandType = xlCmdList
        .Connection = _
        "OLEDB;Provider=Microsoft.Office.List.OLEDB.2.0;Data Source="""";ApplicationName=Excel;Version=12.0.0.0"
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With wb.Connections("Connection")
        .Name = "Connection"
        .Description = ""
    End With
    wb.Connections("Connection").Refresh



clipboard.SetText EmpName
clipboard.PutInClipboard

Set filepick = Application.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)

With filepick
        .AllowMultiSelect = False
        .ButtonName = "File Picker"
        .Title = "File Picker"
        .InitialFileName = sPathHRLF
        .Filters.Add "Custom Word Files", "*.doc, *.docx, *.docm"
        .InitialFileName = EmpName & "*.doc*"
        If (.Show > 0) Then
        End If
        If (.SelectedItems.Count > 0) Then
        SelectedFile = .SelectedItems(1)
        Else
        Exit Sub
        End If
End With
On Error Resume Next

Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
Set objWord = CreateObject("Word.Application")
End If
'Set objDoc = objWord.Documents.Open(SelectedFile)
'Else
Set objDoc = objWord.Documents.Open(SelectedFile)
'End If

objWord.Visible = True
Set objDoc = objWord.ActiveDocument
Set objSelection = objWord.Selection

With WSLF
Set cl2 = .Range("A:A").Find("Leaver Notification Form", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
pos = InStr(1, "Leaver Notification Form sent", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
If pos > 0 Then
Set cl2 = .Range("A:A").Find("Leaver Notification Form", After:=WSLF.Cells(CL.Row, "a"), LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
End If
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(0) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

pos = 0

If cl2 Is Nothing Then
With WSLF
Set cl2 = .Range("A:A").Find("LNF", , LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
'text = WSLF.Cells(cl2.Row, "a").Value
'text = "John Stokes - LNF sent to GSC, IT, Pcard and TandE"
pos = InStr(1, WSLF.Cells(cl2.Row, "a").Value, "LNF sent", vbTextCompare)
If pos > 0 Then
Set cl2 = .Range("A:A").Find("LNF", After:=WSLF.Cells(cl2.Row, "a"), LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
End If
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(0) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)
End With
End If


With WSLF
If cl2 Is Nothing Then
arrLabvar(0) = "N/A"
End If
End With


With WSLF
Set cl2 = .Range("A:A").Find("GSC", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(1) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)
arrLabvar(2) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

With WSLF
If cl2 Is Nothing Then
arrLabvar(1) = "N/A"
arrLabvar(2) = "N/A"
End If
End With

With WSLF
Set cl2 = .Range("A:A").Find("TandE", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(3) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

With WSLF
If cl2 Is Nothing Then
arrLabvar(3) = "N/A"
End If
End With

With WSLF
Set cl2 = .Range("A:A").Find("Pcard", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(4) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

With WSLF
If cl2 Is Nothing Then
arrLabvar(4) = "N/A"
End If
End With

With WSLF
Set cl2 = .Range("A:A").Find("Fuelcard", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(5) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

With WSLF
If cl2 Is Nothing Then
arrLabvar(5) = "N/A"
End If
End With

With WSLF
Set cl2 = .Range("A:A").Find("first notice", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(6) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

With WSLF
If cl2 Is Nothing Then
MsgBox ("There is no first notice uploaded")
End If
End With

arrLabvar(7) = MadeLeaver
arrLabvar(8) = "N/A"
arrLabvar(9) = "N/A"
pos = 0

With WSLF
Set cl2 = .Range("A:A").Find("Exit interview", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
pos = InStr(1, "has no Exit interview form confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
pos2 = InStr(1, "no Exit interview form confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
pos3 = InStr(1, "no EIF confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
pos4 = InStr(1, "has no EIF confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
If pos > 0 Or pos2 > 0 Or pos3 > 0 Or pos4 > 0 Then
Set cl2 = .Range("A:A").Find("Exit interview", After:=WSLF.Cells(cl2.Row, "a"), LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
End If
If Not cl2 Is Nothing Then
arrLabvar(10) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)
End If
End With

pos = 0

With WSLF
If cl2 Is Nothing Then
Set cl2 = .Range("A:A").Find("EIF", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
pos = InStr(1, "has no Exit interview form confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
pos2 = InStr(1, "no Exit interview form confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
pos3 = InStr(1, "no EIF confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
pos4 = InStr(1, "has no EIF confirmation", WSLF.Cells(cl2.Row, "a"), vbTextCompare)
If pos > 0 Or pos2 > 0 Or pos3 > 0 Or pos4 > 0 Then
Set cl2 = .Range("A:A").Find("EIF", After:=WSLF.Cells(cl2.Row, "a"), LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
End If
If Not cl2 Is Nothing Then
arrLabvar(10) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)
End If
End If
End With

With WSLF
If cl2 Is Nothing Then
arrLabvar(10) = "N/A"
End If
End With

With WSLF
Set cl2 = .Range("A:A").Find("Alphabet", LookIn:=xlValues, MatchCase:=False, SearchFormat:=False)
If Not cl2 Is Nothing Then cl2.Select
arrLabvar(11) = Split(WSLF.Cells(cl2.Row, "b"), " ")(0)

End With

With WSLF
If cl2 Is Nothing Then
arrLabvar(11) = "N/A"
End If
End With

arrTxt = Array("Copy of completed Leaver Notification Form placed on personal file ", "Copy of notification to GSC placed on personal file", _
    "Copy of notification to GB-IT Procurement placed on personal file", "Copy of notification to T&E Process Coordinator placed on personal file", _
    "Copy of notification to P-card Helpdesk placed on personal file", "Copy of notification to Fuel Card contact placed on personal file", _
    "Copy of notification to SAP workflow contact placed on personal file  ", "Make a leaver in payroll system", "Close position in SAP ", _
    "Time Administrator/Supervisor details updated in SAP (if appropriate)", "Exit Interview Form received and placed on file (if appropriate)", _
    "Complete Starter/Leaver Notification form and email to Alphabet")

With objSelection.Find
        .MatchWholeWord = True
        .MatchCase = False
        .text = "Copy of notification to IT Procurement placed on personal file"
        .Execute

    If .Found = True Then
    GoTo Replace
        Else: GoTo NoReplace
    End If
End With

Replace:

With objSelection.endof
objSelection.Tables(1).Select
objSelection.Tables(1).Delete
Set HRDoc = objWord.Documents.Open("C:\In Progress\Leaver Process\Leaver Forms\Appendix 2 Employee Leaver Process HR Leaver Form.docx")
Set HRSelection = objWord.Selection
End With

With HRSelection.Find
.MatchWholeWord = True
.MatchCase = False
.text = "Copy of notification to GB-IT Procurement placed on personal file"
.Execute
    If .Found = True Then
        HRSelection.endof
        HRSelection.Tables(1).Select
        HRSelection.Copy
        objWord.ActiveDocument.Close 'SaveChanges:=wdDoNotSaveChanges
        Set HRDoc = Nothing
    End If
End With

Set objSelection = objWord.Selection
objSelection.PasteSpecial
Set HRDoc = Nothing

objSelection.endof

NoReplace:

z = 0
For Each vTxt In arrTxt
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
objSelection.MoveUp Count:=1
'objSelection.HomeKey Unit:=wdStory, Extend:=wdExtend
With objSelection.Find
    .MatchWholeWord = True
    .MatchCase = False
    .text = "Section 3"
    .Execute
End With
objSelection.endof

With objSelection.Find
    .MatchWholeWord = True
    .MatchCase = False
    .text = vTxt
    .Execute
End With

objSelection.endof
objSelection.MoveRight Count:=1

If Not arrLabvar(z) = "" Then
    objSelection.text = arrLabvar(z)
'Else
'    objSelection.Text = "N/A"
End If

objSelection.endof
z = z + 1

Next vTxt

With objSelection.Find
    .MatchWholeWord = True
    .MatchCase = False
    .text = "HR Administrator Name"
    .Execute
End With

objSelection.endof
objSelection.MoveRight Count:=1
objSelection.text = "Jozsef Rado"
objSelection.endof

objSelection.MoveUp Count:=25

With objSelection.Find
    .MatchWholeWord = True
    .MatchCase = False
    .text = "Date filed"
    .Execute
End With

objSelection.endof
objSelection.MoveUp Count:=2
objSelection.MoveLeft
objSelection.MoveUp
objSelection.text = "Jozsef Rado"
objSelection.MoveRight Count:=2
objSelection.text = "Mate Erdei"


        'objSelection.HomeKey Unit:=wdStory
'    With objSelection.Find
'        .MatchWholeWord = True
'        .MatchCase = False
'        .Text = "Calculated"
'        .Execute
'    End With
'        objSelection.endof
'        objSelection.MoveDown Count:=2
'        objSelection.Text = "Jozsef Rado"
'        objSelection.endof
'        objSelection.MoveRight Count:=1
'        objSelection.Text = "Mate Erdei"

If objDoc = (EmpName & " HR Leaver Form.docx") Then
    objDoc.SaveAs ("C:\In Progress\" & EmpName & " HR Leaver Form FINAL.docx")
    End If
If objDoc = (EmpName & " HR Leaver Form.doc") Then
    objDoc.SaveAs ("C:\In Progress\" & EmpName & " HR Leaver Form FINAL.doc")
End If
If objDoc = (EmpName & " HR Leaver Form.docm") Then
    objDoc.SaveAs ("C:\In Progress\" & EmpName & " HR Leaver Form FINAL.docm")
End If

Set objDoc = Nothing
Set objWord = Nothing

End Sub

Private Sub CommandButton7_Click()

Dim wsml As Worksheet, WSCHP As Worksheet, WSED As Worksheet
Dim rngSel As Range, LeavingR As String
Dim FR As Integer, LR As Integer, RAND As Integer
Dim EmpNum As String, EmpName As String, EmpFirstName As String, posnum As String, EmpGrade As String
Dim ManName As String, RedLeaveD As String
Dim ManFirstName As String
Dim EntDate As String, dEntDate As Date, LeaveD As String, LeaveDD As Date, x As Integer, y As Integer, TodaysM As String, LeavingM As String, b As Integer, Date1 As Date, Date2 As Date
Dim StartDateD As Date, StartDateS As String, wyear As Integer
Dim FirstDay As String, LastDay As String
Dim Salary As String, Salary2 As Double, TakenHol As Double, sTakenHol As String, TakenSapHol As Double, AnnualEnt As Double, wmonths As Integer
Dim SAProw As Double, SAPcol As Double, iSAPfield As Double, strSAPfield As Double, NotFree As Integer, LastDay2 As Date, z As Double, e As Double, Free As Integer
Dim outs1 As Double, outs2 As Double, wdays As Double, HolPay As Double, CalMethod As String, SendtoP As String
Dim Holanswer As Variant, strPath As String, FileToOpen As String
Dim CL As Range, q As Integer, StartM As Integer, SAPwsMonth As String, SAPwsYear As String
Dim SapGuiAuto
Dim SAPApp
Dim SAPCon
Dim session
Dim iWeekWdays1 As Integer, iWeekWdays2 As Integer, iDivNum As Double


Dim EmpFirstN As String, EmpLastN As String
Dim objSelection
Dim objWord As Word.Application
Dim objDoc As Word.Document

Set wsml = Application.ThisWorkbook.Worksheets("Managing Leavers")
Set rngSel = Selection

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window) on that connection

EmpNum = wsml.Cells(ActiveCell.Row, "b")
FR = rngSel.Row
LR = rngSel.Find("*", After:=rngSel.Cells(1), LookIn:=xlValues, SearchDirection:=xlPrevious).Row

Call CommandButton18_Click

'WSML.Range("B" & FR & ":E" & LR).Copy WSCHP.Range("A" & WSCHP.Cells(WSCHP.Rows.Count, 1).End(xlUp).Row + 1)
'WSCHP.Activate

With wsml.Cells
        Set CL = .Find(EmpNum, LookIn:=xlValues)
End With

EmpName = wsml.Cells(CL.Row, "c")

LeaveD = wsml.Cells(CL.Row, "e")

LeavingR = wsml.Cells(CL.Row, "d")


If LeavingR = "Redundancy" Then
RedLeaveD = InputBox("The leaving reason is Redundancy, please provide the advised date for the holiday payment calculation.")
RedLeaveD = Replace(RedLeaveD, ".", "/")
End If

''session.findById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
''session.findById("wnd[0]").sendVKey 0
''session.findById("wnd[0]/usr/ctxtRP50G-PERNR").text = Empnum
''session.findById("wnd[0]").sendVKey 0
''session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "41"
''session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").caretPosition = 2
''session.findById("wnd[0]").sendVKey 0
''session.findById("wnd[0]/tbar[1]/btn[7]").press
''StartDateS = session.findById("wnd[0]/usr/ctxtP0041-DAT01").text
''session.findById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
''session.findById("wnd[0]").sendVKey 0

'StartDateS = WSED.Cells(cl.Row, "k")
'StartDateS = Replace(StartDateS, ".", "/")
'StartDateD = DateSerial(Split(StartDateS, "/")(2), Split(StartDateS, "/")(1), Split(StartDateS, "/")(0))
'
'If Year(LeaveD) - Year(StartDateD) > 2 Then
'AnnualEnt = 25
'GoTo SAPEntCheck
'End If
'If (Year(LeaveD)) - (Year(StartDateD)) = -2 Then
'
'        ElseIf Day(StartDateD) = 1 Then AnnualEnt = 25
'            ElseIf Month(LeaveD) = "12" And Day(LeaveD) = "31" Then AnnualEnt = 25
'Else: AnnualEnt = 23
'GoTo SAPEntCheck
'End If
'If Year(LeaveD) - Year(StartDateD) < 2 Then
'    AnnualEnt = 23
'GoTo SAPEntCheck
'End If


'Problem : ElseIf Day(StartDateD) = 1











SAPEntCheck:

'With WSCHP.Cells
'        Set cl = .Find(EmpNum, LookIn:=xlValues)
'        If Not cl Is Nothing Then
'            WSCHP.Activate
'            cl.Select
'        End If
'End With

'If WSCHP.Cells(cl.Row, "f") = "" Then
On Error Resume Next
sTakenHol = InputBox("Please give the taken holidays of " & EmpName)
If sTakenHol = "" Then Exit Sub
TakenHol = CDbl(sTakenHol)

'On Error Resume Next
'session.findById("wnd[0]/tbar[0]/okcd").text = "pa20"
'session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
'
'If Err.Number <> 0 Then
'    session.findById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
'    session.findById("wnd[0]").sendVKey 0
'    session.findById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
'End If
'
'session.findById("wnd[0]").sendVKey 0

session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "8"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 7
session.FindById("wnd[0]/usr/subSUBSCREEN_TC0008:MP000800:0300/tblMP000800TC/txtQ0008-BETRG[4,0]").SetFocus
Salary = session.FindById("wnd[0]/usr/txtQ0008-ANSAL").text
session.FindById("wnd[0]").SendVKey 3

'WSCHP.Cells(cl.Row, "e") = Salary
Salary = Replace(Salary, ",", "")
Salary2 = CDbl(Salary)


session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "2006"
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").SetFocus
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").caretPosition = 4
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330/radRP50G-TIMRD").SetFocus
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330/radRP50G-TIMRD").Select
session.FindById("wnd[0]/tbar[1]/btn[20]").Press


For x = 0 To 100

EntDate = session.FindById("wnd[0]/usr/tblMP200000TC3000/txtP2006-BEGDA[0," & x & "]").text
If Split(EntDate, ".")(2) = Split(LeaveD, "/")(2) Then
    If session.FindById("wnd[0]/usr/tblMP200000TC3000/txtT556B-KTEXT[3," & x & "]").text = "Annual Leave" Then GoTo NExtStep
End If
Next x

NExtStep:

Date1 = CDate(Replace(EntDate, ".", "/"))
                    Date1 = Format(Date1, "dd mm yyyy")
                    Date2 = LeaveD

                    If LeavingR = "Redundancy" Then
                    Date2 = RedLeaveD
                    Else
                    Date2 = LeaveD
                    End If

                    Date2 = Format(Date2, "dd mm yyyy")
                    FirstDay = DateSerial(Year(Date1), Month(Date1), 1)
                    LastDay = DateSerial(Year(Date2), Month(Date2) + 1, 1 - 1)
                    wmonths = DateDiff("m", Date1, Date2)

                    If Date1 = FirstDay And Date2 = LastDay Then
                        wmonths = wmonths + 1
                                ElseIf Date1 <> FirstDay And Date2 <> LastDay Then
                                wmonths = wmonths - 1
                    End If

                    If wmonths < 0 Then wmonths = 0

AnnualEnt = session.FindById("wnd[0]/usr/tblMP200000TC3000/txtP2006-ANZHL[6," & x & "]").text
session.FindById("wnd[0]").SendVKey 3

If AnnualEnt = 23 Or AnnualEnt = 25 Then
AnnualEnt = AnnualEnt
Else
    If Day(Date1) = 1 And Month(Date1) = 1 Then
    AnnualEnt = AnnualEnt
    Else
    AnnualEnt = 23
    End If
End If


session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").text = "7"
        session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").caretPosition = 1
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]").SendVKey 7
        session.FindById("wnd[0]").SendVKey 31

        TodaysM = Format(Date, "m")
        LeavingM = Format(LeaveD, "m")


        LeaveDD = DateSerial(Year(LeaveD), Month(LeaveD), Day(LeaveD))

        If TodaysM = LeavingM And Year(Date) = Year(LeaveD) Then
        Else

            If LeaveDD > Date Then
                Do
                session.FindById("wnd[0]").SendVKey 19
                SAPwsMonth = session.FindById("wnd[0]/usr/txtT247-LTX").text
                SAPwsYear = session.FindById("wnd[0]/usr/txtT552A-KJAHR").text
                Loop Until SAPwsYear = Year(LeaveD) And Month(DateValue("01 " & SAPwsMonth & " 1900")) = LeavingM
            End If

            If LeaveDD < Date Then
                Do
                session.FindById("wnd[0]").SendVKey 18
                SAPwsMonth = session.FindById("wnd[0]/usr/txtT247-LTX").text
                SAPwsYear = session.FindById("wnd[0]/usr/txtT552A-KJAHR").text
                Loop Until SAPwsYear = Year(LeaveD) And Month(DateValue("01 " & SAPwsMonth & " 1900")) = LeavingM
            End If

        End If


                SAProw = 1
Next_SAProw:
                Select Case SAProw

                Case "1"

                                    For iSAPfield = 1 To 7

                                    strSAPfield = CStr(iSAPfield)




                                Select Case strSAPfield
                                        Case "1"
                                            SAPcol = 7

                                        Case "2"
                                            SAPcol = 17

                                        Case "3"
                                            SAPcol = 29

                                        Case "4"
                                            SAPcol = 40

                                        Case "5"
                                            SAPcol = 51

                                        Case "6"
                                            SAPcol = 61

                                        Case "7"
                                            SAPcol = 73

                                End Select

                                            On Error Resume Next


                                            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                                                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                Free = Free + 1

                                                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                    NotFree = NotFree + 1
                                                    Free = 0
                                            End If



Check_Next_SAP_Field:

                                     Next iSAPfield


                Case 3

                        For iSAPfield = 1 To 7

                                    strSAPfield = CStr(iSAPfield)




                                Select Case strSAPfield
                                        Case "1"
                                            SAPcol = 7

                                        Case "2"
                                            SAPcol = 17

                                        Case "3"
                                            SAPcol = 29

                                        Case "4"
                                            SAPcol = 40

                                        Case "5"
                                            SAPcol = 51

                                        Case "6"
                                            SAPcol = 61

                                        Case "7"
                                            SAPcol = 73

                                End Select

                                            On Error Resume Next

                                            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                                                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                Free = Free + 1

                                                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                    NotFree = NotFree + 1
                                                    Free = 0

                                            End If

                                    Next iSAPfield

                 Case 5

                        For iSAPfield = 1 To 7

                                    strSAPfield = CStr(iSAPfield)




                                Select Case strSAPfield
                                        Case "1"
                                            SAPcol = 7

                                        Case "2"
                                            SAPcol = 17

                                        Case "3"
                                            SAPcol = 29

                                        Case "4"
                                            SAPcol = 40

                                        Case "5"
                                            SAPcol = 51

                                        Case "6"
                                            SAPcol = 61

                                        Case "7"
                                            SAPcol = 73

                                End Select

                                            On Error Resume Next

                                            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                                                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                Free = Free + 1

                                                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                    NotFree = NotFree + 1
                                                    Free = 0
                                            End If

                                    Next iSAPfield

                 Case 7

                        For iSAPfield = 1 To 7

                                    strSAPfield = CStr(iSAPfield)




                                Select Case strSAPfield
                                        Case "1"
                                            SAPcol = 7

                                        Case "2"
                                            SAPcol = 17

                                        Case "3"
                                            SAPcol = 29

                                        Case "4"
                                            SAPcol = 40

                                        Case "5"
                                            SAPcol = 51

                                        Case "6"
                                            SAPcol = 61

                                        Case "7"
                                            SAPcol = 73

                                End Select

                                            On Error Resume Next

                                            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                                                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                Free = Free + 1

                                                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                    NotFree = NotFree + 1
                                                    Free = 0
                                            End If

                                    Next iSAPfield

                 Case 9

                        For iSAPfield = 1 To 7

                                    strSAPfield = CStr(iSAPfield)




                                Select Case strSAPfield
                                        Case "1"
                                            SAPcol = 7

                                        Case "2"
                                            SAPcol = 17

                                        Case "3"
                                            SAPcol = 29

                                        Case "4"
                                            SAPcol = 40

                                        Case "5"
                                            SAPcol = 51

                                        Case "6"
                                            SAPcol = 61

                                        Case "7"
                                            SAPcol = 73

                                End Select

                                            On Error Resume Next

                                            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                                                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                Free = Free + 1

                                                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                    NotFree = NotFree + 1
                                                    Free = 0

                                            End If

                                    Next iSAPfield

                    Case 11

                        For iSAPfield = 1 To 7

                                    strSAPfield = CStr(iSAPfield)




                                Select Case strSAPfield
                                        Case "1"
                                            SAPcol = 7

                                        Case "2"
                                            SAPcol = 17

                                        Case "3"
                                            SAPcol = 29

                                        Case "4"
                                            SAPcol = 40

                                        Case "5"
                                            SAPcol = 51

                                        Case "6"
                                            SAPcol = 61

                                        Case "7"
                                            SAPcol = 73

                                End Select

                                            On Error Resume Next

                                            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                                                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                Free = Free + 1

                                                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                                                    NotFree = NotFree + 1
                                                    Free = 0

                                            End If

                                    Next iSAPfield


             End Select
                    If SAProw = 1 Then
                    SAProw = SAProw + 2
                    NotFree = 0
                    GoTo Next_SAProw
                    End If
                    
                    If SAProw = 3 Then
                    iWeekWdays1 = NotFree
                    SAProw = SAProw + 2
                    NotFree = 0
                    GoTo Next_SAProw
                    End If
                    
                    If SAProw = 5 Then
                        iWeekWdays2 = NotFree
                    End If
                    
                    If Not iWeekWdays1 = iWeekWdays2 Then
                        If iWeekWdays2 > iWeekWdays1 Then
                        iWeekWdays1 = iWeekWdays2
                        End If
                    End If
                    
                    Select Case iWeekWdays1
                        Case 7
                        iDivNum = 1 / 260
                        Case 6
                        iDivNum = 1 / 260
                        Case 5
                        iDivNum = 1 / 260
                        Case 4
                        iDivNum = 1 / 208
                        Case 3
                        iDivNum = 1 / 156
                        Case 2
                        iDivNum = 1 / 104
                        End Select

                    

                           wdays = NotFree

                            LastDay2 = DateSerial(Year(LeaveD), Month(LeaveD) + 1, 1 - 1)
                            'LeaveDD = DateSerial(Year(LeaveD), Month(LeaveD), Day(LeaveD))
                            z = LastDay2 - LeaveDD
                            If z = 0 Then
                            ElseIf z <= Free Then
                            wmonths = wmonths + 1
                            End If
                                session.FindById("wnd[0]/tbar[0]/btn[3]").Press
                                'session.findById("wnd[0]/tbar[0]/btn[3]").press
                                'session.findById("wnd[0]/tbar[0]/btn[3]").press



If Not EntDate = "" Then StartM = Month(CDate(Replace(EntDate, ".", "/")))

If Not EntDate = "01.01." & Split(LeaveD, "/")(2) Then
session.FindById("wnd[0]").SendVKey 31

If StartM = Month(Date) Then
    Else
    Do
    session.FindById("wnd[0]").SendVKey 18
    SAPwsMonth = session.FindById("wnd[0]/usr/txtT247-LTX").text
    SAPwsYear = session.FindById("wnd[0]/usr/txtT552A-KJAHR").text
    Loop Until SAPwsYear = Year(CDate(Replace(EntDate, ".", "/"))) And Month(DateValue("01 " & SAPwsMonth & " 1900")) = StartM
End If


'megn?zi a nem teljes kezd? h?nap is teljesnek sz?m?t e
SAProw = 1
Free = 0

    For iSAPfield = 1 To 7

    strSAPfield = CStr(iSAPfield)




Select Case strSAPfield
        Case "1"
            SAPcol = 7

        Case "2"
            SAPcol = 17

        Case "3"
            SAPcol = 29

        Case "4"
            SAPcol = 40

        Case "5"
            SAPcol = 51

        Case "6"
            SAPcol = 61

        Case "7"
            SAPcol = 73

End Select

            On Error Resume Next


            If session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "" Then

                ElseIf session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                Free = Free + 1

                    ElseIf Not session.FindById("wnd[0]/usr/sub:SAPMP51S:1010/ctxtPSHFT-TPR0" & strSAPfield & "[" & SAProw & "," & SAPcol & "]").text = "FREE" Then
                    NotFree = NotFree + 1
                    GoTo CheckDays
            End If


     Next iSAPfield

CheckDays:
If Free = Day(CDate(Replace(EntDate, ".", "/"))) - 1 Then
wmonths = wmonths + 1
End If

End If

session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press



'sz?mol?s
        e = AnnualEnt / 12 * wmonths

        outs1 = Round(e + 0.001)
        If e - outs1 = -0.5 Then
                    outs1 = e
        ElseIf e - outs1 > 0 Then
                    outs1 = outs1 + 0.5

        Else: outs1 = outs1

        End If
        outs2 = outs1 - TakenHol

        'wdays = WSCHP.Cells(cl.Row, "i")
        HolPay = Salary * (iDivNum) * outs2
        HolPay = Round(HolPay, 2)

       If HolPay > 0 Then

                        CalMethod = "Holidays: " & AnnualEnt & "/ " & "12 " & "* " & wmonths & " = " & e & vbNewLine & _
                        outs1 & " - " & TakenHol & " = " & outs2 & vbNewLine & _
                        Salary & " * 1/260 * " & outs2 & " = " & HolPay & " GBP"
                        'WSCHP.Cells(cl.Row, "k") = CalMethod

                        SendtoP = "Hello Edit," & vbNewLine & vbNewLine & "Please arrange " & HolPay & " GBP holiday payment for the following employee:" & vbNewLine & "-" & EmpNum & vbNewLine & "-" & EmpName & vbNewLine & "-" & LeaveD & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"
                        UserForm5.TextBox1 = EmpName & " - Holiday Payment"
                        UserForm5.TextBox2 = SendtoP & vbNewLine & vbNewLine & CalMethod
                        UserForm5.Show
'                        WSCHP.Cells(cl.Row, "l") = HolPay & " GBP"
'                        WSCHP.Cells(cl.Row, "m") = EmpName & " - Holiday Payment" & vbNewLine & SendtoP


                     ElseIf wmonths = 0 And TakenHol = 0 Then

                        CalMethod = "Holidays: " & AnnualEnt & "/ " & "12 " & "* " & wmonths & " = " & e & vbNewLine & _
                        "Not entitled to holiday payment"
                        MsgBox "The employee is not entitled to holiday payment"
'                        WSCHP.Cells(cl.Row, "k") = CalMethod
'                        WSCHP.Cells(cl.Row, "l") = "Not entitled to holiday payment"

                     ElseIf outs1 - TakenHol = 0 Then

                        CalMethod = "Holidays: " & AnnualEnt & "/ " & "12 " & "* " & wmonths & " = " & e & vbNewLine & _
                        outs1 & " - " & TakenHol & " = " & outs2 & vbNewLine & _
                        "Full entitlement taken"
                        MsgBox ("Full entitlement taken")
'                        WSCHP.Cells(cl.Row, "k") = CalMethod
'                        WSCHP.Cells(cl.Row, "l") = "Full entitlement taken"
                     Else

                        HolPay = Abs(HolPay)

                        CalMethod = "Holidays: " & AnnualEnt & "/ " & "12 " & "* " & wmonths & " = " & e & vbNewLine & _
                        outs1 & " - " & TakenHol & " = " & outs2 & vbNewLine & _
                        Salary & " * 1/260 * " & outs2 & " = " & HolPay & " GBP deduction"
'                        WSCHP.Cells(cl.Row, "k") = CalMethod

                        SendtoP = "Hello Edit," & vbNewLine & vbNewLine & "Please arrange " & HolPay & " GBP holiday deduction for the following employee:" & vbNewLine & "-" & EmpNum & vbNewLine & "-" & EmpName & vbNewLine & "-" & LeaveD & vbNewLine & vbNewLine & "Thank you and kind regards," & vbNewLine & "Joe"

'                        ManName = WSED.Cells(CL.Row, "e")
'                        ManFirstName = Split(ManName)(0)
'                        EmpFirstName = Split(EmpName)(0)

                        UserForm5.TextBox1 = EmpName & " - Holiday Deduction"
                        UserForm5.TextBox2 = SendtoP & vbNewLine & vbNewLine & CalMethod
                        UserForm5.Show

                        UserForm6.TextBox1 = ManName
                        UserForm6.TextBox2 = EmpName & " - Holiday Deduction"
                        UserForm6.TextBox3 = "Hello " & ManFirstName & "," & vbNewLine & vbNewLine & "Could you please confirm " & HolPay & " GBP holiday deduction for " & EmpName & " as the number of taken holidays is more than " & EmpFirstName & "'s entitlement (entitled to " & outs1 & ", taken " & TakenHol & ")." & vbNewLine & vbNewLine & _
                        "Thank you and kind regards," & vbNewLine & "Joe"

                        UserForm6.Show

'                        With WSCHP.Cells
'                                Set cl = .Find(Empnum, LookIn:=xlValues)
'                                If Not cl Is Nothing Then
'                                    WSCHP.Activate
'                                    cl.Select
'                                End If
'                        End With
'                        WSCHP.Cells(cl.Row, "l") = HolPay & " GBP DEDUCTION"
'                        WSCHP.Cells(cl.Row, "m") = EmpName & " - Holiday Deduction" & vbNewLine & SendtoP



        End If




'If Not Split(EntDate, ".")(0) = "01" Then
'
'        session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU").getAbsoluteRow(1).Selected = True
'        session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,1]").SetFocus
'        session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,1]").caretPosition = 0
'        session.findById("wnd[0]").sendVKey 7
'        session.findById("wnd[0]").sendVKey 19
'
'        PosNum = session.findById("wnd[0]/usr/ctxtP0001-PLANS").text
'
'        session.findById("wnd[0]").sendVKey 3
'        session.findById("wnd[0]").sendVKey 3
'        session.findById("wnd[0]/tbar[0]/okcd").text = "po13d"
'        session.findById("wnd[0]").sendVKey 0
'        session.findById("wnd[0]/usr/ctxtPM0D1-SEARK").text = PosNum
'        session.findById("wnd[0]").sendVKey 0
'        session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/radPM0D1-TIMRD").Select
'        session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/tblSAPMH5A0INFO_CONTROL").verticalScrollbar.Position = 25
'        session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/tblSAPMH5A0INFO_CONTROL").getAbsoluteRow(27).Selected = True
'        session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/tblSAPMH5A0INFO_CONTROL/txtTT_T777T-ITEXT[0,2]").SetFocus
'        session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/tblSAPMH5A0INFO_CONTROL/txtTT_T777T-ITEXT[0,2]").caretPosition = 0
'        session.findById("wnd[0]").sendVKey 0
'        session.findById("wnd[0]").sendVKey 7
'
'        EmpGrade = session.findById("wnd[0]/usr/txtZCXRHT024-TEXT_BSAL").text
'
'        session.findById("wnd[0]").sendVKey 3
'        session.findById("wnd[0]").sendVKey 3
'End If

'WSCHP.Cells(cl.Row, "g") = AnnualEnt
'WSCHP.Cells(cl.Row, "f") = TakenHol
'WSCHP.Cells(cl.Row, "e") = Salary & " GBP"

With wsml.Cells
        Set CL = .Find(EmpNum, LookIn:=xlValues)
        If Not CL Is Nothing Then
            wsml.Activate
            CL.Select
        End If
End With


Holanswer = MsgBox("Would you like to copy the holiday payment to the HR Leaver Form?", vbYesNo, "Holiday Payment")
If Holanswer = vbYes Then

        EmpFirstN = Split(EmpName, " ")(0)
        EmpLastN = Split(EmpName, " ")(1)

        On Error Resume Next
        Set objWord = GetObject(, "Word.Application")
         'Set objDoc = objWord.Documents.Open("\\Client\V$\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.docm")
        
        Set objDoc = objWord.Documents.Open("C:\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.docx")
        Set objDoc = objWord.Documents.Open("C:\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.doc")
        Set objDoc = objWord.Documents.Open("C:\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.docm")
            If objDoc Is Nothing Then
            Set objWord = CreateObject("Word.Application")

                Set objDoc = objWord.Documents.Open("C:\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.docx")
                Set objDoc = objWord.Documents.Open("C:\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.doc")
                Set objDoc = objWord.Documents.Open("C:\In Progress\" & EmpFirstN & " " & EmpLastN & " HR Leaver Form.docm")
            Else
            End If


            objWord.Visible = True
            Set objSelection = objWord.Selection

                With objSelection.Find
                .MatchWholeWord = True
                .MatchCase = False
                .Execute FindText:="Holidays:"
                .Forward = True
                End With

        objSelection.TypeText (CalMethod)

        Set objWord = Nothing
        Set objDoc = Nothing

End If

wsml.Activate
'WSCHP.Range("a1").CurrentRegion.Sort key1:=WSCHP.Range("d2"), order1:=xlAscending, Header:=xlYes

ThisWorkbook.Save

End Sub

Private Sub CommandButton8_Click()

Dim clipboard As MSForms.DataObject
Dim WSLF As Worksheet, wsml As Worksheet, wb As Workbook, EmpNum As String, EmpName As String, EmpFchar As String, LeaveD As String, LeavingM As String, LeavingY As String, LeaveR As String
Dim CL As Range, Months(1 To 12) As String
Dim sfindVal As String
Dim LR As Integer

Set wb = Application.ThisWorkbook
sfindVal = ActiveCell.Value

Set WSLF = Application.ThisWorkbook.Worksheets("Leaver Folder")
Set wsml = ThisWorkbook.Worksheets("Managing Leavers")
EmpNum = sfindVal

Set clipboard = New MSForms.DataObject

With wsml
Set CL = .Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
If Not CL Is Nothing Then CL.Select
End With

EmpNum = sfindVal

EmpName = wsml.Cells(CL.Row, "c")
EmpFchar = Left(EmpName, 1)
LeaveD = wsml.Cells(CL.Row, "e")
LeavingM = Split(LeaveD, "/")(1)
LeavingY = Split(LeaveD, "/")(2)
LeaveR = wsml.Cells(CL.Row, "d")

Months(1) = "January"
Months(2) = "February"
Months(3) = "March"
Months(4) = "April"
Months(5) = "May"
Months(6) = "June"
Months(7) = "July"
Months(8) = "August"
Months(9) = "September"
Months(10) = "October"
Months(11) = "November"
Months(12) = "December"

    With wb.Connections("Connection").OLEDBConnection
        .BackgroundQuery = False
        .CommandText = Array( _
        "<LIST><VIEWGUID>{42EEE4CF-14EF-4B58-B95D-61924126CBE2}</VIEWGUID><LISTNAME>{B8910762-E3D1-4AA2-A04E-7D516CED1381}</" _
        , _
        "LISTNAME><LISTWEB>http://cdm2.cemex.com/sites/SSC_HRS/HRA_UK/_vti_bin</LISTWEB><LISTSUBWEB></LISTSUBWEB><ROOTFOLDER" _
        , _
        ">/sites/SSC_HRS/HRA_UK/Shared Documents/Employee files/" & EmpFchar & "/" & EmpName & " " & EmpNum & "/Leaver " & Months(LeavingM) & " " & LeavingY & Chr(13) & "" & Chr(10) & "</ROOTFOLDER></LIST>" _
        )
        '.CommandType = xlCmdList
        .Connection = _
        "OLEDB;Provider=Microsoft.Office.List.OLEDB.2.0;Data Source="""";ApplicationName=Excel;Version=12.0.0.0"
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With wb.Connections("Connection")
        .Name = "Connection"
        .Description = ""
    End With
    wb.Connections("Connection").Refresh

LR = WSLF.Cells(Rows.Count, 1).End(xlUp).Row

Const FName As String = "C:\Users\HUZZ01HN\Desktop\VBA\Leaver test\Screenshot.jpg"
Dim pic_rng As Range
Dim ShTemp As Worksheet
Dim ChTemp As Chart
Dim PicTemp As Picture
Application.ScreenUpdating = False
Set pic_rng = WSLF.Range("A1:B" & LR)
Set ShTemp = Worksheets.Add
Charts.Add
ActiveChart.Location Where:=xlLocationAsObject, Name:=ShTemp.Name
Set ChTemp = ActiveChart
pic_rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
ChTemp.Paste
Set PicTemp = Selection
With ChTemp.Parent
.Width = PicTemp.Width + 100
.Height = PicTemp.Height + 150
End With
ChTemp.Export FileName:="C:\Users\HUZZ01HN\Desktop\VBA\Leaver test\Screenshot.jpg", FilterName:="jpg"
Application.DisplayAlerts = False
ShTemp.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True

UserForm7.Image1.Picture = LoadPicture(FName)
UserForm7.Label1.Caption = EmpName & " - " & LeaveR & " - " & LeaveD
UserForm7.Show

End Sub

Private Sub CommandButton9_Click()

Dim posnum As String, LR As Integer, x As Integer, EmpName As String, MsgAns As Integer, MsgAns2 As Integer
Dim CL As Range
Dim wsml As Worksheet
Dim EmpNum As String, LeaveD As Date

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window) on that connection

Set wsml = Application.ThisWorkbook.Worksheets("Managing Leavers")

EmpNum = wsml.Cells(ActiveCell.Row, "b")
Set CL = wsml.Range("B:B").Find(EmpNum, LookIn:=xlValues, SearchFormat:=False)
LeaveD = wsml.Cells(CL.Row, "e")

session.FindById("wnd[0]/tbar[0]/okcd").text = "/n pa20"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtRP50G-PERNR").text = EmpNum
session.FindById("wnd[0]").SendVKey 0
EmpName = session.FindById("wnd[0]/usr/subSUBSCR_HEADER:/1PAPAXX/HDR_50060A:0100/txt$_DG01_500A60_DAT_P0001_ENAME").text
session.FindById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU").getAbsoluteRow(1).Selected = True
session.FindById("wnd[0]").SendVKey 7

If session.FindById("wnd[0]/usr/ctxtP0001-PLANS").text = "99999999" Then
session.FindById("wnd[0]").SendVKey 19
End If

posnum = session.FindById("wnd[0]/usr/ctxtP0001-PLANS").text

MsgAns = MsgBox("Are you sure that you want to delimit " & EmpName & "'s position?", vbYesNo)
    If MsgAns = vbYes Then
        Else
        Exit Sub
    End If
session.FindById("wnd[0]/tbar[0]/okcd").text = "/n po13"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtPM0D1-SEARK").text = posnum
session.FindById("wnd[0]").SendVKey 0
'check if there is an active holder
session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/radPM0D1-TIMRD").Select
session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/tblSAPMH5A0INFO_CONTROL").getAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/radPM0D1-TIMRD").SetFocus
session.FindById("wnd[0]").SendVKey 20
x = 0

Do While Not session.FindById("wnd[0]/usr/tblMP100100TC3000/txtQ1001-VERKTXT[4," & x & "]").text = "__________"

If session.FindById("wnd[0]/usr/tblMP100100TC3000/txtQ1001-VERKTXT[4," & x & "]").text = "Holder" Then
    If session.FindById("wnd[0]/usr/tblMP100100TC3000/txtP1001-ENDDA[1," & x & "]").text = "31.12.9999" Then
    MsgBox ("The position has an active holder!")
    Exit Sub
    End If
End If
x = x + 1
If x = 25 Then
session.FindById("wnd[0]/usr/tblMP100100TC3000").verticalScrollbar.Position = 25
x = 0
End If
Loop
session.FindById("wnd[0]").SendVKey 3






session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/tblSAPMH5A0INFO_CONTROL").getAbsoluteRow(0).Selected = True
session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB01/ssubSUB1:SAPMH5A0:5005/radPM0D1-TIMRD").Select
session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select
session.FindById("wnd[1]/usr/ctxtPM0D1-GREDT1").text = Replace(CStr(DateSerial(Year(LeaveD), Month(LeaveD), Day(LeaveD)) + 1), "/", ".")
session.FindById("wnd[1]").SendVKey 8
session.FindById("wnd[2]/usr/btnSPOP-OPTION1").Press

If session.FindById("wnd[0]/sbar").text = "Object 01 S " & posnum & " is delimited on " & Replace(CStr(DateSerial(Year(LeaveD), Month(LeaveD), Day(LeaveD)) + 1), "/", ".") Then
MsgBox ("The position has been delimited successfully.")
End If


End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

End Sub

