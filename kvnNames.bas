Attribute VB_Name = "kvnNames_v1_2"
' БИБЛИОТЕКА kvnNames - отвязка данных от статичных имен
' Автор KVN
' v. 1.2 от 07.06.2017 г.
' ----------------------------------------------------------------------
' Книга должна содержать лист 'kvnNames' с таблицей из следующих колонок
' Имя, Значение, Адрес, Примечание

Private Const cnstSheet_kvnNames = "kvnNames"

Public Function kvnNames_GetFormula(rng As Range, Optional p_Local As Boolean = False, Optional Пересчёт As Boolean = True) As String
    Application.Volatile Пересчёт
    
    If p_Local Then
        kvnNames_GetFormula = rng.FormulaLocal
    Else
        kvnNames_GetFormula = rng.formula
    End If
End Function
Public Function kvnNames_GetProp(p_PropName)
    On Error Resume Next
    kvnNames_GetProp = 0
    kvnNames_GetProp = Application.WorksheetFunction.VLookup(p_PropName, ActiveWorkbook.Worksheets(cnstSheet_kvnNames).Range("A2:B1000"), 2, False)
    On Error GoTo 0
    If kvnNames_GetProp = 0 Then
        MsgBox ("Свойство " & p_PropName & " не найдено на листе " & cnstSheet_kvnNames)
        MsgBox (1 / 0)
        Exit Function
    End If
End Function
Public Function kvnNames_GetCell(p_Name As String, Optional row As Integer = -9999) As Range
    Dim idx, formula, r
    
    On Error Resume Next
    idx = 0
    idx = Application.WorksheetFunction.Match(p_Name, ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("$A:$A"), 0)
    On Error GoTo 0
    If idx = 0 Then
        MsgBox ("Имя " & p_Name & " не найдено на листе " & cnstSheet_kvnNames)
        MsgBox (1 / 0)
        Exit Function
    End If
    
    formula = Mid(ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("B" & idx).formula, 2, 255)
    If row = -9999 Then r = ActiveCell.row Else r = row
    Set kvnNames_GetCell = Range(formula).Offset(r - Range(formula).row, 0)
End Function

Public Function kvnNames_FindValue(p_Name As String, p_val) As Range
    Dim idx, formula
    Dim rng As Range
    
    On Error Resume Next
    idx = 0
    idx = Application.WorksheetFunction.Match(p_Name, ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("$A:$A"), 0)
    On Error GoTo 0
    If idx = 0 Then
        MsgBox ("Имя " & p_Name & " не найдено на листе " & cnstSheet_kvnNames)
        MsgBox (1 / 0)
        Exit Function
    End If
    
    formula = Mid(ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("B" & idx).formula, 2, 255)
    Set rng = Range(formula)
    
    On Error Resume Next
    idx = 0
    idx = Application.WorksheetFunction.Match(p_val, rng.Worksheet.Columns(rng.Column), 0)
    On Error GoTo 0
    If idx = 0 Then
        MsgBox ("Значение " & p_val & " не найдено в диапазоне '" & rng.Worksheet.Name & "'!" & rng.Worksheet.Columns(rng.Column).Address)
        MsgBox (1 / 0)
        Exit Function
    End If
    
    Set kvnNames_FindValue = rng.Offset(idx - rng.row, 0)
End Function
Public Function kvnNames_GetSheet(p_Name As String) As Worksheet
    Dim idx, formula
    
    On Error Resume Next
    idx = 0
    idx = Application.WorksheetFunction.Match(p_Name, ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("$A:$A"), 0)
    On Error GoTo 0
    If idx = 0 Then
        MsgBox ("Имя " & p_Name & " не найдено на листе " & cnstSheet_kvnNames)
        MsgBox (1 / 0)
        Exit Function
    End If
    
    formula = ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("B" & idx).Value
    Set kvnNames_GetSheet = ActiveWorkbook.Sheets(formula)
End Function
Public Function kvnNames_Match(p_Name As String, p_val) As String
    Dim idx, formula
    Dim rng As Range
    Dim s1, s2
    
    On Error Resume Next
    idx = 0
    idx = Application.WorksheetFunction.Match(p_Name, ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("$A:$A"), 0)
    On Error GoTo 0
    If idx = 0 Then
        MsgBox ("Имя " & p_Name & " не найдено на листе " & cnstSheet_kvnNames)
        MsgBox (1 / 0)
        Exit Function
    End If
    
    formula = Mid(ActiveWorkbook.Sheets(cnstSheet_kvnNames).Range("B" & idx).formula, 2, 255)
    Set rng = Range(formula)
    kvnNames_Match = "=Match(" & p_val & ", '" & rng.Worksheet.Name & "'!" & rng.Worksheet.Columns(rng.Column).Address & ",0)"
End Function
