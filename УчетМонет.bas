Attribute VB_Name = "УчетМонет"
' Модуль учета монет из драгоценных металлов
'
' ------------------------------------------
' (C) KVN v.1.1 от 31.05.2017
' ------------------------------------------

' ***********************************************************************************
' Комментируем нужную строку в зависимости от того, на чьей стороне запускаем макрос.
#Const KIND_OF_INSTANCE = "MAIN"
'#Const KIND_OF_INSTANCE = "DO"
' ***********************************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function WNetGetUserA Lib "mpr.dll" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
    
    Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpszOp As String, _
        ByVal lpszFile As String, ByVal lpszParams As String, _
        ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function WNetGetUserA Lib "mpr.dll" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
    
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpszOp As String, _
        ByVal lpszFile As String, ByVal lpszParams As String, _
        ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
#End If

Public FormResult As Boolean
Function rr_ReplaceWord(pName, pValue)
    ObjWord.Selection.Find.ClearFormatting
    ObjWord.Selection.Find.Replacement.ClearFormatting
    With ObjWord.Selection.Find
        .Text = "$" & pName & "$"
        .Replacement.Text = pValue
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    ObjWord.Selection.Find.Execute Replace:=wdReplaceAll
End Function
Function GetUserName() As String
    Dim sUserNameBuff As String * 255
    sUserNameBuff = Space(255)
    Call WNetGetUserA(vbNullString, sUserNameBuff, 255&)
    GetUserName = Left$(sUserNameBuff, InStr(sUserNameBuff, vbNullChar) - 1)
End Function
'JSON
'{
'   "firstName": "Иван",
'   "lastName": "Иванов",
'   "address": {
'       "streetAddress": "Московское ш., 101, кв.101",
'       "city": "Ленинград",
'       "postalCode": 101101
'   },
'   "phoneNumbers": [
'       "812 123-1234",
'       "916 123-4567"
'   ]
'}
Private Function FormString(c_Name, c_Value, Optional Last As Boolean = False)
    FormString = Chr(9) & Chr(34) & c_Name & Chr(34) & ": " & Chr(34) & c_Value & Chr(34)
    If Not Last Then FormString = FormString & ","
End Function
Function CreateJSON(rw As Integer) As String
    Dim fName, i
    Dim Email_Body
    
    kvnNames_GetCell("В_наличии!Статус", rw).Value = "оформляется"
    
    Email_Body = ""
    
    fName = kvnNames_GetProp("ПутьКФайлуJSON") & "S_" & ActiveWorkbook.Sheets("Настройки").Range("D8").Value & ".json" ' ID монеты
    Open fName For Output As #1
    Print #1, "{"
    For i = 8 To 14
        tmp = FormString(ActiveWorkbook.Sheets("Настройки").Range("A" & i).Value, ActiveWorkbook.Sheets("Настройки").Range("D" & i).Value, i = 14)
        If i <= 12 Then Email_Body = Email_Body & tmp & "%0D%0A"
        Print #1, tmp
    Next i
    Print #1, "}"
    Close #1
    
#If KIND_OF_INSTANCE = "DO" Then
    Email_Subject = "Монета №" & ActiveWorkbook.Sheets("Настройки").Range("D8").Value & " реализована"
    Email_Send_To = kvnNames_GetProp("КомуПосылатьИнфОПродаже")
    
    Email_Body = Replace(Email_Body, Chr(34), "'")
    Mail_Object = "mailto:" & Email_Send_To & "?subject=" & Email_Subject & " &body=" & Email_Body
    'Range("A" & rw & ":F33").Select
    'Selection.Copy
    ShellExecute 0&, vbNullString, Mail_Object, vbNullString, vbNullString, vbNormalFocus
#End If
    
    CreateJSON = "Данные о монете переданы. Файл " & fName
End Function
Sub CopyToSale(rw As Integer)
    Dim off
    
    kvnNames_GetCell("В_наличии!Статус", rw).Value = "продано"
    Range("A" & rw & ":V" & rw).Select
    Selection.Copy
    
    kvnNames_GetSheet("ИмяЛистаПродано").Select                             ' Перешли на лист "Продано"
    Range("A3").Select
    off = 0
    While (ActiveCell.Offset(off, 0).Value <> "")
        off = off + 1
    Wend
    
    ActiveCell.Offset(off, 0).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    off = ActiveCell.row
    Application.CutCopyMode = False
    
    Range("W" & off).Value = ActiveWorkbook.Sheets("Настройки").Range("D12").Value        ' Дата реализации
    Range("X" & off & ":AA" & off).Select
    Selection.FillDown
    
    Range("AB" & off).Value = ActiveWorkbook.Sheets("Настройки").Range("D9").Value        ' Продал
    Range("AC" & off).Value = ActiveWorkbook.Sheets("Настройки").Range("D10").Value       ' Покупатель
    Range("AD" & off).Value = ActiveWorkbook.Sheets("Настройки").Range("D11").Value       ' Контакт
        
    Range("AE" & off).Select
    Selection.FillDown
    
    kvnNames_GetSheet("ИмяЛиста_РаспРеализ").Select
    Range("D1").Value = ActiveWorkbook.Sheets("Настройки").Range("D12").Value             ' Дата
    Range("C5").Value = ActiveWorkbook.Sheets("Настройки").Range("D8").Value              ' ID монеты
    Range("D32").Value = ActiveWorkbook.Sheets("Настройки").Range("D9").Value             ' Продал

End Sub
Sub ПродажаМонеты_Sub(FlagFromJSON As Boolean)
Attribute ПродажаМонеты_Sub.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim idx_Spr_UserName, cUserName, idx, numDO
    Dim idx_SheetName_Retail, idx_Retail_ID, rw As Integer, mes
    Dim r1, r2, CoinID
    
    If FlagFromJSON Then
        kvnNames_GetSheet("ИмяЛиста_В_наличии").Select
        CoinID = ActiveWorkbook.Sheets("Настройки").Range("D8").Value
        kvnNames_FindValue("В_наличии!УникНомерМонеты", CoinID).Select
        rw = ActiveCell.row
    Else
        rw = ActiveCell.row
    End If
    
#If KIND_OF_INSTANCE = "DO" Then
    If ThisWorkbook.Name = ActiveWorkbook.Name Then
        MsgBox ("Макрос запускается на активной книге БД Монет")
        Exit Sub
    End If
#End If
    
    idx_SheetName_Retail = kvnNames_GetProp("ИмяЛиста_В_наличии")
    If ActiveSheet.Name <> idx_SheetName_Retail Then
        MsgBox ("Макрос формирования документов по проданным монетам запускается с листа '" & idx_SheetName_Retail & "'!")
        Exit Sub
    End If
    
    idx_Spr_UserName = kvnNames_GetProp("idx_Spr_UserName")
    cUserName = GetUserName()
    If IsError(Application.Match(cUserName, Range(idx_Spr_UserName), 0)) = True Then
        MsgBox ("К сожалению, у Вас нет прав на формирование документов.")
        Exit Sub
    End If
    
    idx = Application.Match(cUserName, Range(idx_Spr_UserName), 0)
    numDO = Application.WorksheetFunction.Index(Range(kvnNames_GetProp("idx_Spr_DO_for_UserName")), idx, 0)
    
    If kvnNames_GetCell("В_наличии!Статус", rw).Value <> kvnNames_GetProp("Статус_ВНаличии") Then
        MsgBox ("Для запуска нужно стоять на строке с монетой со статусом 'в наличии'!")
        Exit Sub
    End If
    
    If kvnNames_GetCell("В_наличии!МестоХранения", rw).Value = kvnNames_GetProp("НаименованиеХранилища") Then
        MsgBox ("Продажа монеты из хранилища невозможна!")
        Exit Sub
    End If
    
    If numDO <> "Все" Then
        If kvnNames_GetCell("В_наличии!МестоХранения", rw).Value <> numDO Then
            MsgBox ("Вам разрешена только продажа монет подразделения " & numDO & "!")
            Exit Sub
        End If
    End If
    
    ActiveWorkbook.Sheets("Настройки").Range("D3").Value = ActiveCell.row
    ActiveWorkbook.Sheets("Настройки").Range("D14").Value = idx + 1
    
    If Not FlagFromJSON Then
        AktSale.Show
        If Not FormResult Then Exit Sub
    End If
    
#If KIND_OF_INSTANCE = "DO" Then
    mes = CreateJSON(rw)
    MsgBox ("Покупка оформлена." & vbCrLf & mes)
#Else
    Call CopyToSale(rw)
    Call Send_Email_Using_Keys
    Sleep 200
    'MsgBox ("Покупка оформлена")
#End If

End Sub
Private Sub Send_Email_Using_Keys()
    Dim Mail_Object As String
    Dim Email_Subject, Email_Send_To, Email_Cc, Email_Bcc, Email_Body As String
    Dim hwnd As Long, nm
    
    Email_Subject = "Распоряжение (Реализация монеты №" & Range("C5").Value & ")"
    Email_Send_To = kvnNames_GetProp("КомуПосылатьРаспоряжения")
        
    'Email_Cc = "exceltipmail@gmail.com "
    'Email_Bcc = "exceltipmail@gmail.com "
    'Email_Body = "%0D%0A" & Sheets("Настройки").Range("D15").Value & "%0D%0A"
    Mail_Object = "mailto:" & Email_Send_To & "?subject=" & Email_Subject '& " &body=" & Email_Body

    ' *****************************************
    Sheets("Расп реализация").Select
' Sheets("Распоряжение").Copy
'    Cells.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'    Application.CutCopyMode = False
    
    Range("A1:F33").Select
    Selection.Copy
    ' *****************************************

'    On Error GoTo debugs

    ShellExecute 0&, vbNullString, Mail_Object, vbNullString, vbNullString, vbNormalFocus
    'Application.Wait (Now + TimeValue("0:00:03"))
    
    ' SHIFT +
    ' CTRL  ^
    ' ALT   %
    
    'hwnd = GetForegroundWindow
    'nm = GetCaption(hwnd)
    'If InStr(nm, Email_Subject & kvnGetProp("idx_OutlookWindowCaption")) > 0 Then
    '    Application.SendKeys "+({INSERT})"
    '    '   Application.SendKeys "%(ь)"
    '    '   Application.SendKeys ("{ENTER}")
    '    Application.StatusBar = False
    'Else
    '    Application.StatusBar = hwnd & " -> " & nm
    'End If
    
'debugs:
'    If Err.Description <> "" Then MsgBox Err.Description
End Sub
' ****************************************************************************************************************
' Cледующие процедуры (СохранитьБД_Монет, ЗагрузитьДанные_о_ПроданнойМонете, УдалитьПроданнуюМонетуИзНаличных, СкопируемМакросы) доступны только в инстанции Main
' ****************************************************************************************************************
#If KIND_OF_INSTANCE = "MAIN" Then
Sub ОбновитьБД_Монет()
    Dim fName, CurName, vbFileName
    
    fName = kvnNames_GetProp("ПутьКФайлуБазыДанныхДляДО") & kvnNames_GetProp("ИмяФайлаБазыДанныхДляДО")
    CurName = ActiveWorkbook.Name
    
    If Dir(fName) <> "" Then                            ' Файл существует
        SetAttr fName, vbNormal                         ' Снимем атрибут Read Only
        Kill (fName)                                    ' Удалим файл
    End If

    Windows(ThisWorkbook.Name).Activate                 ' Скопируем данные с текущей книги в новую
    Sheets(Array("Выполнение плана", "В наличии", "Технический лист", "Настройки", "kvnNames")).Copy
    kvnNames_GetSheet("ИмяЛиста_В_наличии").Columns("S:X").EntireColumn.Hidden = True ' Скроем колонки с S по X
    Range("A1").Select
    
    ' Вставить данные по выполнению плана, как значения
    kvnNames_GetSheet("ИмяЛиста_ВыполнениеПлана").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Сохранить файл
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=fName, FileFormat:=xlExcel8, CreateBackup:=False 'xlOpenXMLWorkbook
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    SetAttr fName, vbReadOnly
    
    Windows(CurName).Activate
    kvnNames_GetSheet("ИмяЛиста_ВыполнениеПлана").Select
    
    MsgBox ("Книга " & fName & " обновлена")
End Sub

' *.JSON
' -----------------------------------------
'{
'    "ID_монеты": "1347",
'    "Продал": "Козулин В.Н.",
'    "Клиент": "Петров Петр Петрович",
'    "Контакт": "2-33-223",
'    "ДатаРеализации": "30.05.2017",
'    "Login": "KozulinVN",
'    "User": "50"
'}
' -----------------------------------------
Sub ЗагрузитьДанные_о_ПроданнойМонете()
    Dim fName, fPath, t
    Dim cName, cValue
    Dim PathChange
    
    PathChange = kvnNames_GetProp("ПутьКФайлуJSON")
    
    fPath = PathChange & "S_*.JSON"
    fName = Dir(fPath, vbNormal)
    If (fName = "") Then
        MsgBox ("Активизируйте доп. офисы, ни одной монеты не продано.")
        Exit Sub
    Else
        MsgBox ("Загружаем файл " & fName)
    End If
    
    Open PathChange & fName For Input As #1
    While Not EOF(1)
        Line Input #1, s
        If s <> "{" And s <> "}" Then
            If Right(s, 1) = "," Then s = Left(s, Len(s) - 1)
            t = InStr(s, ":")
            cName = Mid(s, 3, t - 4)
            cValue = Mid(s, t + 3, Len(s) - t - 3)
            '=MATCH(cName,A8:A14,0)
            t = Application.WorksheetFunction.Match(cName, ActiveWorkbook.Sheets("Настройки").Range("A8:A14"), 0)
            If ActiveWorkbook.Sheets("Настройки").Range("D7").Offset(t, 0).NumberFormat = "m/d/yyyy" Then
                cValue = CDate(cValue)
            End If
            ActiveWorkbook.Sheets("Настройки").Range("D7").Offset(t, 0).Value = cValue
        End If
    Wend
    Close #1
    
    Call ПродажаМонеты_Sub(True)
    Kill PathChange & fName
    
    If Dir() <> "" Then
        MsgBox ("В каталоге еще есть файлы для обработки. Запустите макрос повторно.")
    Else
        MsgBox ("Все файлы в каталоге обработаны.")
    End If
End Sub
Sub УдалитьПроданнуюМонетуИзНаличных()
    Dim id, rw As Integer
    
    If ActiveSheet.Name <> kvnNames_GetProp("ИмяЛистаПродано") Then
        MsgBox ("Запускаемся только на листе '" & kvnNames_GetProp("ИмяЛистаПродано") & "'")
        Exit Sub
    End If
    
    If kvnNames_GetCell("Продано!Статус").Value = "Отражено" Then
        MsgBox ("Монета уже отражена в учете")
        Exit Sub
    Else
        kvnNames_GetCell("Продано!Статус").Value = "Отражено"
    End If
    
    id = kvnNames_GetCell("Продано!УникНомерМонеты").Value
    rw = kvnNames_FindValue("В_наличии!УникНомерМонеты", id).row
    
    kvnNames_GetSheet("ИмяЛиста_В_наличии").Activate
    Rows(rw & ":" & rw).Select
    If kvnNames_GetCell("В_наличии!Статус", rw).Value <> "продано" Then
        MsgBox ("Мы пытаемся удалить не проданную монету!")
        Exit Sub
    End If
        
    If MsgBox("Удалим строку " & t, vbYesNo + vbDefaultButton2) = vbYes Then
        Selection.Delete Shift:=xlUp
    End If
    kvnNames_GetSheet("ИмяЛистаПродано").Activate
    
    MsgBox ("Монета " & id & " удалена из монет 'В наличии'")
    
End Sub
Sub СкопируемМакросы()
    Dim mName
    
    mName = kvnNames_GetProp("ИмяФайлаСМакросамиДляВСП")
    
    If ActiveWorkbook.Name <> mName Then
        MsgBox ("Запускаемся только на '" & mName & "'")
        Exit Sub
    End If
    
    ' Удалим все старые макросы
    With ActiveWorkbook.VBProject.VBComponents
         For iCount& = .Count To 1 Step -1
             Set iVBComponent = .Item(iCount&)
             Select Case iVBComponent.Type
                 Case 1 To 3: .Remove iVBComponent
                 Case 100
                 iVBComponent.CodeModule.DeleteLines 1, iVBComponent.CodeModule.CountOfLines
             End Select
        Next
    End With
    ' Скопируем наши макросы
    With ThisWorkbook.VBProject.VBComponents
         For iCount& = .Count To 1 Step -1
            If .Item(iCount&).Type <= 3 Then
                .Item(iCount&).Export (ThisWorkbook.Path & "\tempMod.bas")
                ActiveWorkbook.VBProject.VBComponents.Import (ThisWorkbook.Path & "\tempMod.bas")
                Kill ThisWorkbook.Path & "\tempMod.bas"
                If Dir(ThisWorkbook.Path & "\tempMod.frx") <> "" Then
                    Kill ThisWorkbook.Path & "\tempMod.frx"
                End If
            End If
         Next
    End With
    ActiveWorkbook.VBProject.VBComponents("УчетМонет").CodeModule.ReplaceLine 9, "'#Const KIND_OF_INSTANCE = ""MAIN"""
    ActiveWorkbook.VBProject.VBComponents("УчетМонет").CodeModule.ReplaceLine 10, "#Const KIND_OF_INSTANCE = ""DO"""
    MsgBox ("Макросы скопированы")
End Sub

#End If
' ****************************************************************************************************************
' Предыдущие процедуры (СохранитьБД_Монет, ЗагрузитьДанные_о_ПроданнойМонете, УдалитьПроданнуюМонетуИзНаличных, СкопируемМакросы) доступны только в инстанции Main
' ****************************************************************************************************************

Sub ПродажаМонеты()
#If KIND_OF_INSTANCE = "DO" Then
    Dim i As Integer, flag As Boolean
    Dim ИмяФайлаБазыДанныхДляДО
    
    ИмяФайлаБазыДанныхДляДО = kvnNames_GetProp("ИмяФайлаБазыДанныхДляДО")
    If ThisWorkbook.Name = ActiveWorkbook.Name Then                     ' Макрос запущен операционистом с книги макросов
        i = 1                                                           ' Попытаемся найти книгу базы данных
        flag = False
        While (i <= Application.Workbooks.Count) And (flag = False)
            If Application.Workbooks(i).Name = ИмяФайлаБазыДанныхДляДО Then
                Windows(ИмяФайлаБазыДанныхДляДО).Activate
                kvnNames_GetSheet("ЛистПоУмолчаниюФайлаБазыДанныхДляДО").Select
                flag = True
            Else
                i = i + 1
            End If
        Wend
        If Not flag Then
            Workbooks.Open Filename:=kvnNames_GetProp("ПутьКФайлуБазыДанныхДляДО") & ИмяФайлаБазыДанныхДляДО, UpdateLinks:=0
            kvnNames_GetSheet("ЛистПоУмолчаниюФайлаБазыДанныхДляДО").Select
        End If
    End If
#End If
    Call ПродажаМонеты_Sub(False)
    
End Sub
