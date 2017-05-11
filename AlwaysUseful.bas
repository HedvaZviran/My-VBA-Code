Attribute VB_Name = "AlwaysUseful"
Option Explicit

Sub NowAndStatus()
Attribute NowAndStatus.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' now מאקרו
'
' קיצור מקשים: Ctrl‏+Shift+N
    Dim Srow As Integer: Srow = ActiveCell.Row
    Dim Scol As Integer: Scol = ActiveCell.Column
    Dim I As Integer
    Dim Stomp As Integer
    Dim Status
    
    'Overwrite checks: whether to overwrite cell contents. Will be asked only if the cell isn't empty and does not contain a date
    'or contains a date which is over a month from today - will be sent to AskStomp
    'In all other cases (empty cell or date less than one month from today) will go straight to YesStomp)
    If ActiveCell.Value = "" Then GoTo YesStomp
    If IsDate(ActiveCell.Value) = False Then GoTo AskStomp
    If IsDate(ActiveCell.Value) = True And Abs(ActiveCell.Value - Date) < 30 Then GoTo YesStomp
    GoTo AskStomp
     
AskStomp: 'Asks whether to overwrite cell contents)
    Stomp = MsgBox("דורסת?", vbYesNo)
        If Stomp = vbYes Then GoTo YesStomp
        If Stomp = vbNo Then Exit Sub
          
YesStomp: 'First stage = enters current date and time as value
        ActiveCell.Value = now
'Second stage = only for worksheets of current surveys that need constant surveillance
'Make sure tab color is the same turqoise as the other survey sheets!
        If ActiveWorkbook.Name = "סטטוס.xlsm" And Scol = 3 And ActiveSheet.Tab.ColorIndex = 42 Then
        'Entering the formula to check urgency
            ActiveCell.Offset(0, 1) = "=IF(INT(NOW()-C" & Srow & ")>$E$1,$E$1,INT(NOW()-C" & Srow & "))"
EnterStatus:
'Updating cell color for status
' 0 in order not to change the color, 1-4 if you want to choose
            Status = InputBox("Input Status:" & vbNewLine & "1=Action required" & vbNewLine & "2=Waiting for answer" & vbNewLine & "3=Audit sent" & vbNewLine & "4=Audit received" & vbNewLine & "0 to exit")
            Select Case Status
                Case 1
                    With Selection.Offset(0, -1).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                
                Case 2
                    With Selection.Offset(0, -1).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                Case 3
                    With Selection.Offset(0, -1).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 15261367
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    
                 Case 4
                    With Selection.Offset(0, -1).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 5296274
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                
                Case 0
                    Exit Sub
                
                Case Else
                    MsgBox ("Status must be 0, 1, 2, 3 or 4")
                    GoTo EnterStatus
            End Select
            
            Call CreateAlphabetizedList
            Call ReArrange
            Cells(Srow, Scol).Activate
            End If

End Sub
Private Sub CreateAlphabetizedList()
Dim I As Integer
Dim LastRow As Integer
Columns("F:G").Delete (xlToLeft)

LastRow = FindLastRow(Cells(1, 1))
   Cells(1, 6) = LastRow - 1
    For I = 2 To LastRow
        Cells(I, 6) = Cells(I, 1)
        If I > 2 Then
            If WorksheetFunction.CountIf(Range(Cells(2, 6), Cells(I - 1, 6)), Cells(I, 6).Value) >= 1 Then
                Cells(I, 7) = "Double"
            End If
        End If
        Cells(I, 2).Copy
        Cells(I, 6).PasteSpecial xlPasteFormats
    Next I
    Columns("F:G").Select
    Application.CutCopyMode = False
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("F2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range(Cells(2, 6), Cells(LastRow, 7))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Columns("F:F").EntireColumn.AutoFit
End Sub
Sub Scroll()
'
' Scroll מאקרו
'
' קיצור מקשים: Ctrl‏+Shift+S
'
    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
    End With
End Sub

Private Sub ReArrange()
'
' Macro3 מאקרו
'
' קיצור מקשים: Ctrl‏+Shift+R
'status: 1=red, 2=yellow, 3=blue 3=green
'Make sure of workbook + sorting red-yellow-blue-green

    Columns("A:D").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add(Range("B2:B200"), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0 _
        , 0)
    ActiveSheet.Sort.SortFields.Add(Range("B2:B200"), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, _
        255, 0)
    ActiveSheet.Sort.SortFields.Add(Range("B2:B200"), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(183, _
        222, 232)
    ActiveSheet.Sort.SortFields.Add Key:=Range("C2:C200") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:C200")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Apply conditional formatting to all of column D
Columns("D:D").Select
    Selection.FormatConditions.Delete
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = "=$E$1"
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueFormula
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = _
        "=$E$1/2"
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 0
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
'Remove conditional formatting from statuses 3, 4 (no more actions required)
Cells(2, 2).Activate
Do Until ActiveCell.Interior.ColorIndex = 24 Or ActiveCell.Row = 300
    ActiveCell.Offset(1, 0).Activate
Loop
Do Until ActiveCell.Interior.Color = RGB(255, 255, 255) Or ActiveCell.Row = 300
    ActiveCell.Offset(0, 2).FormatConditions.Delete
    ActiveCell.Offset(1, 0).Activate
Loop

End Sub

Sub אימייל()
'
' אימייל מאקרו
'
' קיצור מקשים: Ctrl‏+Shift+M

Dim Original As String
Dim Name As String
Dim Email As String
Dim X As Integer
Original = ActiveCell.Value
X = InStr(1, Original, "<")
Name = Trim(Left(Original, X - 1))
Email = Trim(Mid(Original, X + 1, Len(Original) - X - 1))
ActiveCell.Value = Name
ActiveCell.Offset(1, 0) = Email

End Sub

Sub IndustryMacro()
'
' IndustryMacro מאקרו
'
    On Error Resume Next
    SaveSheetCopy
    Scroll
    Rows("2:4").Select
    Selection.EntireRow.Hidden = False
    If IsEmpty(Cells(3, 2)) Then
        Range("A3:AE3").Select
        Selection.Delete Shift:=xlUp
        Rows("3:4").EntireRow.AutoFit
    End If
    FixTitles
    ClearBottom
    
    
End Sub

Sub SaveSheetCopy()
Dim ASName As String
    ASName = ActiveSheet.Name
    Sheets(ASName).Copy Before:=Sheets(1)
    Sheets(1).Name = "מקור " & ASName
    Sheets(ASName).Activate
End Sub

Sub FixTitles()
Dim TitleRow As Integer
Select Case ActiveSheet.Name
Case "טופס נתונים"
    TitleRow = 2
Case "בדיקה טכנית"
    TitleRow = 2
Case "השוואה חיצונית"
    TitleRow = 3
Case Else
    TitleRow = InputBox("באיזו שורה הכותרות?")
End Select

    Cells(TitleRow, 2) = "מספר עובד"
    Cells(TitleRow, 3) = "קוד עיסוק"
    Cells(TitleRow, 4) = "שם עיסוק"
    Cells(TitleRow, 5) = "איזור גיאוגרפי"
    Cells(TitleRow, 6) = "תאריך תחילת עבודה בחברה"
    Cells(TitleRow, 7) = "אחוז משרה"
    Cells(TitleRow, 8) = "רכב חברה (0/1/2)"
    Cells(TitleRow, 9) = "שווי השימוש ברכב (לרכב חברה)"
    Cells(TitleRow, 10) = "שכר בסיסי"
    Cells(TitleRow, 11) = "תשלומים נלווים וגילומים אחרים"
    Cells(TitleRow, 12) = "גילום רכב"
    Cells(TitleRow, 13) = "פרמיות/ תמריצים חודשיים"
    Cells(TitleRow, 14) = "אחוז פרמיות/ תמריץ להפרשה"
    Cells(TitleRow, 15) = "ממוצע חודשי של תשלום בגין שעות נוספות בפועל"
    Cells(TitleRow, 16) = "ממוצע חודשי של תשלום בגין משמרות וכוננויות"
    Cells(TitleRow, 17) = "עמלת מטרה (Target) שנתי מצטברת בגין מכירה (לאנשי מכירות) "
    Cells(TitleRow, 18) = "עמלה בפועל (Actual) שנתית מצטברת בגין מכירה (לאנשי מכירות) "
    Cells(TitleRow, 19) = "אחוז עמלה להפרשה"
    Cells(TitleRow, 20) = "רישוי וביטוח"
    Cells(TitleRow, 21) = "משכורת י" & Chr(34) & "ג"
    Cells(TitleRow, 22) = "אחוז משכורת י" & Chr(34) & "ג להפרשה לסוציאליות"
    Cells(TitleRow, 23) = "בונוס שנתי- Target"
    Cells(TitleRow, 24) = "בונוס שנתי- Actual"
    Cells(TitleRow, 25) = "אחוז בונוס להפרשה"
    Cells(TitleRow, 26) = "דמי הבראה ותשלומים אחרים"
    Cells(TitleRow, 27) = "הפרשת המעביד לקה" & Chr(34) & "ש (%)"
    Cells(TitleRow, 28) = "תקרה לקה" & Chr(34) & "ש (0/1/2)"
    Cells(TitleRow, 29) = "משתנה1"
    Cells(TitleRow, 30) = "משתנה2"
    Cells(TitleRow, 31) = "משתנה3"
    
End Sub

Sub ClearBottom()
    Rows(FindLastRow(Cells(2, 10)) + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete
    Columns(32).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete
    
    Cells(1, 1).Activate
End Sub




Sub PasswordToggle()
Dim Pass As String

If Len(ActiveWorkbook.Password) > 1 Then
    ActiveWorkbook.Password = ""
Else
    Pass = InputBox("Type new password (for default, leave empty)")
    If Pass = "" Then
        ActiveWorkbook.Password = "ZVIRANDATA17"
    Else
        ActiveWorkbook.Password = Pass
    End If
End If
ActiveWorkbook.Save
End Sub

Function FindLastRow(StartPoint As Range)
If IsEmpty(StartPoint) = False Then
    If IsEmpty(StartPoint.Offset(1, 0)) = True Then
        FindLastRow = StartPoint.Row
    Else
        FindLastRow = StartPoint.End(xlDown).Row
    End If
Else
    FindLastRow = StartPoint.End(xlUp).Row
End If
End Function


Sub AddPassword()
Dim DefaultPass
DefaultPass = MsgBox("Use default password?", vbYesNo)
If DefaultPass = vbYes Then
    ActiveWorkbook.Password = "ZVIRANDATA17"
Else
    ActiveWorkbook.Password = InputBox("What password would you like?")
End If
ActiveWorkbook.Close Savechanges:=True
End Sub

Function ModVLookup(LookupValue As String, LookupArray As Range, MatchNumber As Integer, ReturnColumn As Integer)
'vlookup by kth match
Dim NumOfRows As Double
    NumOfRows = LookupArray.Rows

Dim ReturnValue

Dim I As Integer
    
Dim MainLookupCounter As Integer
    MainLookupCounter = 0

For I = LookupArray.Row To NumOfRows
    If Cells(I, LookupArray.Column) = LookupValue Then
        MainLookupCounter = MainLookupCounter + 1
        If MainLookupCounter = MatchNumber Then
            ReturnValue = Cells(I, ReturnColumn)
'            GoTo ValueFound
        End If
    End If
Next I
'GoTo ValueNotFound

'ValueFound:
    ModVLookup = ReturnValue
  '  Exit Function

'ValueNotFound:
  '  If MainLookupCounter = 0 Then
  '      ModVLookup = "LookupValue Not Found"
  '  Else
  '      ModVLookup = "LookupValue appears less times than MatchNumber"
  '  End If

End Function

Function FindByTwo(FirstIdent As String, ScndIdent As String, Rng As Range, FirstCol As Integer, ScndCol As Integer, ResultCol As Integer)
'Vlookup by two identifiers
Dim I As Integer, Combo As String
Combo = FirstIdent & ScndIdent
I = 1
Do Until Rng(I, FirstCol).Value & Rng(I, ScndCol).Value = Combo Or I > Rng.Rows.Count
    I = I + 1
Loop
FindByTwo = Rng(I, ResultCol)
End Function

Function pull(xref As String) As Variant
  'inspired by Bob Phillips and Laurent Longre
  'but written by Harlan Grove
  '-----------------------------------------------------------------
  'Copyright (c) 2003 Harlan Grove.
  '
  'This code is free software; you can redistribute it and/or modify
  'it under the terms of the GNU General Public License as published
  'by the Free Software Foundation; either version 2 of the License,
  'or (at your option) any later version.
  '-----------------------------------------------------------------
  '2004-05-30
  'still more fixes, this time to address apparent differences between
  'XL8/97 and later versions. Specifically, fixed the InStrRev call,
  'which is fubar in later versions and was using my own hacked version
  'under XL8/97 which was using the wrong argument syntax. Also either
  'XL8/97 didn't choke on CStr(pull) called when pull referred to an
  'array while later versions do, or I never tested the 2004-03-25 fix
  'against multiple cell references.
  '-----------------------------------------------------------------

  '2004-05-28
  'fixed the previous fix - replaced all instances of 'expr' with 'xref'
  'also now checking for initial single quote in xref, and if found
  'advancing past it to get the full pathname [dumb, really dumb!]
  '-----------------------------------------------------------------
  '2004-03-25
  'revised to check if filename in xref exists - if it does, proceed;
  'otherwise, return a #REF! error immediately - this avoids Excel
  'displaying dialogs when the referenced file doesn't exist
  '-----------------------------------------------------------------
  Dim xlapp As Object, xlwb As Workbook
  Dim b As String, R As Range, C As Range, n As Long
  '** begin 2004-05-30 changes **


  '** begin 2004-05-28 changes **
  '** begin 2004-03-25 changes **
  n = InStrRev(xref, "\")
  If n > 0 Then
    If Mid(xref, n, 2) = "\[" Then
      b = Left(xref, n)
      n = InStr(n + 2, xref, "]") - n - 2
      If n > 0 Then b = b & Mid(xref, Len(b) + 2, n)
    Else
      n = InStrRev(Len(xref), xref, "!")
      If n > 0 Then b = Left(xref, n - 1)
    End If

    '** key 2004-05-28 addition **
    If Left(b, 1) = "'" Then b = Mid(b, 2)
    On Error Resume Next
    If n > 0 Then If Dir(b) = "" Then n = 0
    Err.Clear
    On Error GoTo 0
  End If

  If n <= 0 Then
    pull = CVErr(xlErrRef)
    Exit Function
  End If
  '** end 2004-03-25 changes **
  '** end 2004-05-28 changes **
  pull = Evaluate(xref)

  '** key 2004-05-30 addition **
  If IsArray(pull) Then Exit Function
  '** end 2004-05-30 changes **

  If CStr(pull) = CStr(CVErr(xlErrRef)) Then
    On Error GoTo CleanUp   'immediate clean-up at this point

    Set xlapp = CreateObject("Excel.Application")
    Set xlwb = xlapp.Workbooks.Add  'needed by .ExecuteExcel4Macro

    On Error Resume Next    'now clean-up can wait

    n = InStr(InStr(1, xref, "]") + 1, xref, "!")
    b = Mid(xref, 1, n)

    Set R = xlwb.Sheets(1).Range(Mid(xref, n + 1))

    If R Is Nothing Then
      pull = xlapp.ExecuteExcel4Macro(xref)

    Else
      For Each C In R
        C.Value = xlapp.ExecuteExcel4Macro(b & C.Address(1, 1, xlR1C1))
      Next C

      pull = R.Value

    End If

CleanUp:
    If Not xlwb Is Nothing Then xlwb.Close 0
    If Not xlapp Is Nothing Then xlapp.Quit
    Set xlapp = Nothing

  End If

End Function

Function FindRate(PV As Double, FV As Double, Nper As Integer)
FindRate = (FV / PV) ^ (1 / Nper) - 1
End Function
Sub GetWorksheetNames()
Dim I As Integer
For I = 1 To Worksheets.Count
    ActiveCell.Offset(I, 0) = Worksheets(I).Name
Next I
End Sub
Sub GotoAudit()
    If Evaluate("=isref('Audit'!$A$1)") = True Then
        Worksheets("Audit").Activate
    ElseIf Evaluate("=isref('השוואה חיצונית'!$A$1)") = True Then
        Worksheets("השוואה חיצונית").Activate
    ElseIf Evaluate("=isref('בדיקה טכנית'!$A$1)") = True Then
        Worksheets("בדיקה טכנית").Activate
    End If
End Sub

Sub HTMacro()
Call SaveSheetCopy
Call ClearBottom
Cells(1, 1).Activate
End Sub

Sub CellsToString()
Attribute CellsToString.VB_Description = "מעתיק את תוכן התאים הגלויים והופך לרשימה מופרדת על ידי פסיקים שאפשר להדביק."
Attribute CellsToString.VB_ProcData.VB_Invoke_Func = "C\n14"
Dim iRow As Integer
Dim Str As String
Dim iArea As Range
Dim obj As New DataObject

For Each iArea In Selection.Areas
    For iRow = 1 To iArea.Rows.Count
        If iArea.Rows(iRow).Hidden = False Then
            If Len(Str) < 1 Then
                Str = iArea(iRow, 1).Value
            Else
                Str = Str & ", " & iArea(iRow, 1).Value
            End If
        End If
    Next iRow
Next iArea
'Make object's text equal above string variable
  obj.SetText Str

'Place DataObject's text into the Clipboard
  obj.PutInClipboard

End Sub
Sub EnglishTemplate()
Call SaveSheetCopy
Call FixTitles
Dim MsgBoxResult
If ActiveSheet.Name = "Data Template" Then
    ActiveSheet.Name = "טופס נתונים"
ElseIf Left(ActiveSheet.Name, 8) = "External" Then
    ActiveSheet.Name = "השוואה חיצונית"
Else
    MsgBoxResult = MsgBox("Yes עבור טופס נתונים" & vbNewLine & "No עבור השוואה חיצונית", vbYesNoCancel)
    If MsgBoxResult = vbYes Then
        ActiveSheet.Name = "טופס נתונים"
    ElseIf MsgBoxResult = vbNo Then
        ActiveSheet.Name = "השוואה חיצונית"
    Else
        Exit Sub
    End If
End If
      
    
Columns("E:E").Select
    Selection.Replace What:="Center", Replacement:="מרכז", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
         
    Selection.Replace What:="North", Replacement:="צפון", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        Selection.Replace What:="South", Replacement:="דרום", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        Selection.Replace What:="Jerusalem", Replacement:="ירושלים", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        Selection.Replace What:="Haifa/Yokneam/Caesarea", Replacement:="קיסריה/חיפה/יקנעם", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        Selection.Validation.Delete
        
        ActiveSheet.DisplayRightToLeft = True
        
        Call ClearBottom
        Cells(1, 1).Activate
End Sub
Sub AF()
Range("B2:AC2").AutoFilter visibledropdown:=True
End Sub
Sub Testing()
ActiveCell.Interior.Color = 15261367
End Sub

Function EmptyRange(Rng As Range) As Boolean
Dim R As Range
Dim Flag As Boolean
Flag = True
For Each R In Rng
    If IsEmpty(R) = False Then
        Flag = False
        Exit For
    End If
Next R
EmptyRange = Flag
End Function

