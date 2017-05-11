Attribute VB_Name = "ValidationAndLocking"
'������ ��� �� �������� ������ ������ ��� ������ ��� ���������.
'Val1 �-Val2 �� ���� ������ (������ ������ ������, ����� ������ �����)
'Oprtr - ��� ������ ������ ����� (���, ���� �... ���')
'ErrMsg - ����� ������
'ValType - ��� ������ (�����, ����...)
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

'���� ����� ���� �� ��� �������� ������ ������ ����� �����
'������ �� ���� �������� ���, �������� �� ������

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
'���� �� �� ����� ��� ��� ����� ������ ����� ����� ����� ����� ���
'��� ���� - ���� ������� ������ ����� ���� �� ������
'���� ����� ���� ����� �-MarkValidation, ��� �� ��� ��� ������ �� ������
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
'����� �� ������� �� ��� ������ ��� ����� ����� ���� ������� ��������.
'���� ������ �� ReplaceValidation �������� �� �� ���

Val1 = ActiveCell.Validation.Formula1
Call ValToClipboard(Val1)
Val2 = ActiveCell.Validation.Formula2
Oprtr = ActiveCell.Validation.Operator
ErrMsg = ActiveCell.Validation.ErrorMessage
ValType = ActiveCell.Validation.Type

End Sub

Sub ReplaceValidation()
'����� �� ������ ��� ����� (��� �� ���� ���� �� ��)
'������ ������ ���� �� �-CopyValidation ��� ����� ����� ������� ���������
'��� ����, ���� ������ �� ������ ��� ������ ������ �� ������� ���������
'������� ��� ����� �� ���� ����� ������ �� ��� ������

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

'���� ���� ��� �� ���� ������ ������� �� �� ��� ����� ���� ����� ����
'���� ���� �� ��� ����

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
'���� �� �� ����� ��� ��� ����� ������ ����� ������
'���� �� ����� ����� ������ �����, ���� �� ����� ����� ������ �����, ���� ������ �����
'(����� �� �� ���� �� ��� ������ ����� ��� ���� ������ ����� �� ����. ���� ����)

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
'����� ��� ������ (��� �� ����� ����� ����� ����)
'���� ������ ����� ����� �� CTRL+v (����� �����) ��� ��� ����� ���� �����

On Error Resume Next
ActiveCell.PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ActiveCell.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

End Sub



Sub CopyErrMsg()
'����� �� ����� ������ ������ ������ (���� �� ���� ���� ����, �� ������ �� ����� ���)
ErrMsg = ActiveCell.Validation.ErrorMessage
End Sub

Sub ReplaceErrMsg()
'����� �� ����� ������ ��� ����� ����� ����� ������ ������ ������ �������.
'�� ���� �� ���� �� ����� ��� ��� ����� ������ ����� ����� ����
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
'����� �� ��� ����� ������� ��� ����� ������ ������
'Decimal = 2
'List = 3
ValType = ActiveCell.Validation.Type

End Sub

Function TranslateErrmsg(HebErrMes As String)
'�������� �� ����� ������� ���� ����� ������ �� ����� ������
'���� ���� �� ��� ����� ����� ������ ����� ������ ������ ������
On Error Resume Next
If InStr(1, ErrMsg, "�") > 0 Then
    If InStr(1, HebErrMes, "�" & Chr(34) & "�") = 0 Then
        HebErrMes = Replace(HebErrMes, "��� ���� ���� ���", "Please enter value between")
        HebErrMes = Replace(HebErrMes, "�-", "and ")
        HebErrMes = Replace(HebErrMes, "����", "and")
        HebErrMes = Replace(HebErrMes, "���", "days")
        HebErrMes = Replace(HebErrMes, "����", "years")
        HebErrMes = Replace(HebErrMes, "������", "months")
        HebErrMes = Replace(HebErrMes, "����", "days")
    Else
        HebErrMes = Replace(HebErrMes, "��� ���� ���� ���", "Please enter value between ILS")
        HebErrMes = Replace(HebErrMes, "�-", "and ILS ")
        HebErrMes = Replace(HebErrMes, "�" & Chr(34) & "�", "")
    End If
End If
   TranslateErrmsg = HebErrMes
End Function

Public Sub LockAllSheets()
'���� �� �� �������� ������
Dim iWS As Integer

For iWS = 1 To Worksheets.Count
    Worksheets(iWS).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="ZBS"
    Cells(4, 1).Activate
    
Next iWS
End Sub

Public Sub LockSheet()
'
'���� �� ������ �����
'���� ����� �� ������ ������ ������

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="ZBS"
End Sub

Public Sub UnlockSheet()
'����� ����� ����� ����

   ActiveSheet.Unprotect Password:="ZBS"
End Sub

Public Sub UnlockCells()
'
' ���� �����, ����� ���� �������, ���� �����
Dim YN
    Call UnlockSheet
    Selection.Locked = False
    Selection.FormulaHidden = False
    YN = MsgBox("Clear contents?", vbYesNo)
    If YN = vbYes Then Selection.ClearContents
    Call LockSheet
End Sub

Public Sub LockCells()

'����� �����, ���� ���� �������, ���� �����
    Call UnlockSheet
    Selection.Locked = True
    Selection.FormulaHidden = False
    Call LockSheet
End Sub

Public Sub UnlockAllSheets()
'����� ����� �� �� ��������
Dim iWS As Integer

For iWS = 1 To Worksheets.Count
    Worksheets(iWS).Unprotect Password:="ZBS"
    
    
Next iWS
End Sub

Sub HideComments()
'����� �� �� ������ �����
On Error Resume Next
Dim Rng As Range
For Each Rng In Selection
    Rng.Comment.Visible = False
 Next Rng
End Sub


Sub CheckUnlockedNotEmpty()
'���� �� �� ����� ������ ��� ��������, ����� ���� �� ����� ��� ������ ��� �� ����� �� ��.
'�� ��,���� ��� ����� �� ���

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
'���� ����� ����� ������ ���� ����� ���� (�� �� �� �� ������� ��� �������)
Call CopyValidation

If ValType <> xlValidateDecimal And ValType <> xlValidateWholeNumber Then
    MsgBox ("�� ��� ����� ������� �����")
    Exit Sub
End If
    
With ActiveCell.Validation
            .ErrorMessage = "��� ���� ���� ��� " & Val1 & " ���� " & Val2
End With

End Sub

Sub ClearForm()
'����� �� �� ����� ��� ������ ������
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
'����� �� ������ ������ ��� ������ ����� (���� ����� ����� �����)
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
