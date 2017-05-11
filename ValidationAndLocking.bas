Attribute VB_Name = "ValidationAndLocking"
'משתנים שהם לא ספציפיים למאקרו מסויים אלא זמינים לכל המאקרואים.
'Val1 ו-Val2 הם ערכי האימות (גבולות באימות מספרים, רשימה באימות רשימה)
'Oprtr - סוג המגבלה באימות מספרי (בין, גדול מ... וכו')
'ErrMsg - התראת השגיאה
'ValType - סוג האימות (רשימה, מספר...)
Option Explicit
Public Val1 As String
Public Val2 As String
Public Oprtr As XlFormatConditionOperator
Public ErrMsg As String
Public ValType As XlDVType
Public CopiedRange As Range



Sub ChangeSelectedValidationItem()
'
' Ctrl+Shift+O
'

'כאשר רוצים לתקן את אחת האופציות מאימות נתונים בצורת רשימה
'בוחרים את הערך הרלוונטי בתא, ומפעילים את המאקרו

Dim currentvalidation As String
Dim CurrentItem As String
Dim EditedItem As String
GetValType
If ValType <> xlValidateList Then
    MsgBox ("Must be list validation")
    Exit Sub
End If
currentvalidation = ActiveCell.Validation.Formula1
CurrentItem = ActiveCell.Value
EditedItem = InputBox("What is the correct form?", Default:=CurrentItem)
If IsNull(EditedItem) Then Exit Sub
Selection.Cells(1.1) = EditedItem

currentvalidation = Replace(currentvalidation, CurrentItem, EditedItem)

    With ActiveCell.SpecialCells(xlCellTypeSameValidation).Validation
        .Modify Type:=xlValidateList, Formula1:=currentvalidation
    End With
If ActiveCell.Address = ActiveCell.SpecialCells(xlCellTypeSameValidation).Address Then Exit Sub
    With ActiveCell.SpecialCells(xlCellTypeSameValidation).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub ValidatedToWhite()
'הופך את כל התאים שיש בהם אימות נתונים בטווח הנבחר לבעלי מילוי לבן
'ועל הדרך - קובע שבנעילת הגליון התאים הללו לא יינעלו
'משמש בעיקר לאחר שימוש ב-MarkValidation, אבל גם בכל פעם שמשנים את האימות
On Error Resume Next
Dim Rng As Range
Dim v As Integer
v = 0
For Each Rng In Selection
    v = Rng.SpecialCells(xlCellTypeSameValidation).Count
    If v >= 1 Then
        With Rng.SpecialCells(xlCellTypeSameValidation).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16777215
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
        Rng.Locked = False
        Rng.FormulaHidden = False
        v = 0
    End If
Next Rng

End Sub

Sub CopyValidation()
'מעתיק את הנתונים של סוג האימות בתא הפעיל ושומר אותם במשתנים גלובליים.
'משמש בשילוב עם ReplaceValidation שמפעילים על תא אחר

Val1 = ActiveCell.Validation.Formula1
Call ValToClipboard(Val1)
Val2 = ActiveCell.Validation.Formula2
Oprtr = ActiveCell.Validation.Operator
ErrMsg = ActiveCell.Validation.ErrorMessage
ValType = ActiveCell.Validation.Type

End Sub

Sub ReplaceValidation()
'מחליף את האימות בתא הפעיל (בין אם קיים ובין אם לא)
'חייבים להשתמש לפני כן ב-CopyValidation כדי שיהיו ערכים במשתנים הגלובליים
'בכל מקרה, בשלב הראשון של המאקרו הזה מוצגים הערכים של המשתנים הגלובליים
'והאימות בתא יוחלף רק אחרי אישור הערכים על ידי המשתמש

Dim YN As VbMsgBoxResult

If ActiveCell.SpecialCells(xlCellTypeSameValidation).Count > 0 Then
    YN = MsgBox(vbTab & "Type" & vbTab & "Operater" & vbTab & "Val1" & vbTab & "Val2" & vbTab & "Error Message" & vbNewLine _
    & "New" & vbTab & ValType & vbTab & Oprtr & vbTab & Val1 & vbTab & Val2 & vbTab & ErrMsg & vbNewLine _
    & "Old" & vbTab & ActiveCell.Validation.Type & vbTab & ActiveCell.Validation.Operator & vbTab & ActiveCell.Validation.Formula1 _
     & vbTab & ActiveCell.Validation.Formula2 & vbTab & ActiveCell.Validation.ErrorMessage, vbOKCancel)
    If YN = vbCancel Then Exit Sub
    If YN = vbAbort Then Exit Sub
End If

Dim SingleCell As Boolean
SingleCell = False
If Selection.Rows.Count + Selection.Columns.Count = 2 Then SingleCell = True

    With ActiveCell.SpecialCells(xlCellTypeSameValidation).Validation
        .Modify Type:=ValType, AlertStyle:=xlValidAlertStop, Operator:= _
        Oprtr, Formula1:=Val1, Formula2:=Val2
    End With

ReplaceErrMsg
If SingleCell = True Then ValidatedToWhite

End Sub

Sub ChangeErrorMessage()

'פותח תיבת קלט עם הערת השגיאה הנוכחית של תא כדי שניתן יהיה לערוך אותה
'בסוף צובע את התא בלבן

Dim CurrentItem As String
Dim EditedItem As String

CurrentItem = ActiveCell.Validation.ErrorMessage
EditedItem = InputBox("What is the correct error message?", Default:=CurrentItem)
If EditedItem = "" Then Exit Sub

    With ActiveCell.Validation
            .ErrorMessage = EditedItem
    End With

ValidatedToWhite
End Sub

Public Sub MarkValidation()
'מסמן את כל התאים שיש בהם אימות נתונים בטווח המסומן
'תאים עם אימות רשימה נצבעים בכתום, תאים עם התראת שגיאה נצבעים באדום, ובלי נצבעים בכתום
'(במקור כי זה שימש אך ורק לאימות רשימה בלי הערה ואימות מספרי עם הערה. כדאי לשפר)

On Error Resume Next
Dim Rng As Range
Dim v As Integer
Dim Col As Long
v = 0
For Each Rng In Selection
    v = Rng.SpecialCells(xlCellTypeSameValidation).Count
    If v >= 1 Then
        If Rng.Validation.Type = 3 Then
            Col = 49407
        Else
            Col = 255
            Rng.Value = Rng.Validation.Formula1
        End If
        With Rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = Col
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
        v = 0
    End If
Next Rng
End Sub

Sub PasteValueNotes()
'מדביק ערך והערות (כדי לא לדרוס עיצוב מותנה קיים)
'אצלי הגדרתי קיצור מקלדת של CTRL+v (הדבקה רגילה) כדי שלא אדביק רגיל בטעות

On Error Resume Next
ActiveCell.PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ActiveCell.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

End Sub



Sub CopyErrMsg()
'מעתיק את התראת השגיאה למשתנה גלובלי (בעצם לא עושה כלום בלבד, רק בשילוב עם מאקרו אחר)
ErrMsg = ActiveCell.Validation.ErrorMessage
End Sub

Sub ReplaceErrMsg()
'מחליף את התראת השגיאה בכל התאים בטווח הנבחר בהתראה שנשמרה במשתנה הגלובלי.
'על הדרך גם צובע את התאים שיש בהם אימות נתונים בטווח הנבחר בלבן
Dim R As Range
Dim Rng As Range
Dim Trans
Set R = Selection
Trans = MsgBox("Translate?", vbYesNo)
If Trans = vbCancel Then Exit Sub
If Trans = vbYes Then ErrMsg = TranslateErrmsg(ErrMsg)
For Each Rng In R
With Rng.Validation
    .ErrorMessage = ErrMsg
End With
ValidatedToWhite
Next Rng
End Sub

Sub GetValType()
'מעתיק את סוג אימות הנתונים בתא הפעיל למשתנה גלובלי
'Decimal = 2
'List = 3
ValType = ActiveCell.Validation.Type

End Sub

Function TranslateErrmsg(HebErrMes As String)
'פונקצייה של תרגום אוטומטי פשוט ומאוד ספציפי של התראת השגיאה
'כדאי לשפר על ידי שימוש בערכי האימות במקום בהודעת השגיאה הקיימת
On Error Resume Next
If InStr(1, ErrMsg, "א") > 0 Then
    If InStr(1, HebErrMes, "ש" & Chr(34) & "ח") = 0 Then
        HebErrMes = Replace(HebErrMes, "אנא ציין נתון בין", "Please enter value between")
        HebErrMes = Replace(HebErrMes, "ל-", "and ")
        HebErrMes = Replace(HebErrMes, "לבין", "and")
        HebErrMes = Replace(HebErrMes, "יום", "days")
        HebErrMes = Replace(HebErrMes, "שנים", "years")
        HebErrMes = Replace(HebErrMes, "חודשים", "months")
        HebErrMes = Replace(HebErrMes, "ימים", "days")
    Else
        HebErrMes = Replace(HebErrMes, "אנא ציין נתון בין", "Please enter value between ILS")
        HebErrMes = Replace(HebErrMes, "ל-", "and ILS ")
        HebErrMes = Replace(HebErrMes, "ש" & Chr(34) & "ח", "")
    End If
End If
   TranslateErrmsg = HebErrMes
End Function

Public Sub LockAllSheets()
'נועל את כל הגליונות בחוברת
Dim iWS As Integer

For iWS = 1 To Worksheets.Count
    Worksheets(iWS).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="ZBS"
    Cells(4, 1).Activate
    
Next iWS
End Sub

Public Sub LockSheet()
'
'נועל את הגליון הפעיל
'כדאי להפוך את הסיסמא למשתנה גלובלי

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="ZBS"
End Sub

Public Sub UnlockSheet()
'משחרר נעילת גליון פעיל

   ActiveSheet.Unprotect Password:="ZBS"
End Sub

Public Sub UnlockCells()
'
' מסיר נעילה, משחרר תאים מסומנים, נועל גליון
Dim YN
    Call UnlockSheet
    Selection.Locked = False
    Selection.FormulaHidden = False
    YN = MsgBox("Clear contents?", vbYesNo)
    If YN = vbYes Then Selection.ClearContents
    Call LockSheet
End Sub

Public Sub LockCells()

'משחרר נעילה, נועל תאים מסומנים, נועל גליון
    Call UnlockSheet
    Selection.Locked = True
    Selection.FormulaHidden = False
    Call LockSheet
End Sub

Public Sub UnlockAllSheets()
'משחרר נעילה של כל הגליונות
Dim iWS As Integer

For iWS = 1 To Worksheets.Count
    Worksheets(iWS).Unprotect Password:="ZBS"
    
    
Next iWS
End Sub

Sub HideComments()
'מסתיר את כל ההערות בטווח
On Error Resume Next
Dim Rng As Range
For Each Rng In Selection
    Rng.Comment.Visible = False
 Next Rng
End Sub


Sub CheckUnlockedNotEmpty()
'עובר על כל התאים בשימוש בכל הגליונות, ובודק עבור כל התאים הלא נעולים האם הם ריקים או לא.
'אם לא,ישאל האם לרוקן את התא

Dim iWS As Worksheet
Dim Rng As Range
Dim Clear
For Each iWS In Worksheets
    iWS.Activate
    For Each Rng In iWS.UsedRange
        If Rng.Locked = False Then
            If Rng.Value <> "" Then
                Clear = MsgBox(iWS.Name & " " & Rng.Address & " " & Rng.Value & vbNewLine & "Clear?", vbYesNo)
                If Clear = vbYes Then
                    If Rng.MergeCells = True Then
                        Rng.MergeArea.ClearContents
                    Else
                        Rng.ClearContents
                    End If
                End If
            End If
        End If
    Next Rng
Next iWS
End Sub

Sub AutoValidateError()
'יוצר התראת אימות נתונים עבור אימות טווח (רק אם יש גם מינימום וגם מקסימום)
Call CopyValidation

If ValType <> xlValidateDecimal And ValType <> xlValidateWholeNumber Then
    MsgBox ("לא סוג אימות הנתונים הנכון")
    Exit Sub
End If
    
With ActiveCell.Validation
            .ErrorMessage = "אנא ציין נתון בין " & Val1 & " לבין " & Val2
End With

End Sub

Sub ClearForm()
'מרוקן את כל התאים שהם למילוי המשתמש
Dim Rng As Range
For Each Rng In ActiveSheet.UsedRange
If Rng.Locked = False And Rng.Font.Color = 0 Then
    If Rng.MergeCells = True Then
        Rng.MergeArea.ClearContents
    Else
        Rng.ClearContents
    End If
End If
Next Rng
End Sub

Sub ClearConditionalFormatting()
'מנטרל את העיצוב המותנה מכל הגליון הפעיל (כולל פתיחת נעילת גליון)
    On Error Resume Next
    Dim Rng As Range
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
        WS.UsedRange.FormatConditions.Delete
          
        
        For Each Rng In WS.UsedRange
            Rng.Comment.Visible = True
        Next Rng
    Next WS
End Sub

Sub ValToClipboard(Str As String)
Dim obj As New DataObject
'Make object's text equal above string variable
  obj.SetText Str

'Place DataObject's text into the Clipboard
  obj.PutInClipboard
End Sub

Sub CopyComment()
Call ValToClipboard(ActiveCell.Comment.Text)
End Sub

Function LeftMinusOne(Rng As Range)
LeftMinusOne = Left(Rng.Value, Len(Rng.Value) - 1)
End Function
