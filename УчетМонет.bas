Attribute VB_Name = "���������"
' ������ ����� ����� �� ����������� ��������
'
' ------------------------------------------
' (C) KVN v.1.1 �� 31.05.2017
' ------------------------------------------

' ***********************************************************************************
' ������������ ������ ������ � ����������� �� ����, �� ���� ������� ��������� ������.
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
'   "firstName": "����",
'   "lastName": "������",
'   "address": {
'       "streetAddress": "���������� �., 101, ��.101",
'       "city": "���������",
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
    
    kvnNames_GetCell("�_�������!������", rw).Value = "�����������"
    
    Email_Body = ""
    
    fName = kvnNames_GetProp("����������JSON") & "S_" & ActiveWorkbook.Sheets("���������").Range("D8").Value & ".json" ' ID ������
    Open fName For Output As #1
    Print #1, "{"
    For i = 8 To 14
        tmp = FormString(ActiveWorkbook.Sheets("���������").Range("A" & i).Value, ActiveWorkbook.Sheets("���������").Range("D" & i).Value, i = 14)
        If i <= 12 Then Email_Body = Email_Body & tmp & "%0D%0A"
        Print #1, tmp
    Next i
    Print #1, "}"
    Close #1
    
#If KIND_OF_INSTANCE = "DO" Then
    Email_Subject = "������ �" & ActiveWorkbook.Sheets("���������").Range("D8").Value & " �����������"
    Email_Send_To = kvnNames_GetProp("�����������������������")
    
    Email_Body = Replace(Email_Body, Chr(34), "'")
    Mail_Object = "mailto:" & Email_Send_To & "?subject=" & Email_Subject & " &body=" & Email_Body
    'Range("A" & rw & ":F33").Select
    'Selection.Copy
    ShellExecute 0&, vbNullString, Mail_Object, vbNullString, vbNullString, vbNormalFocus
#End If
    
    CreateJSON = "������ � ������ ��������. ���� " & fName
End Function
Sub CopyToSale(rw As Integer)
    Dim off
    
    kvnNames_GetCell("�_�������!������", rw).Value = "�������"
    Range("A" & rw & ":V" & rw).Select
    Selection.Copy
    
    kvnNames_GetSheet("���������������").Select                             ' ������� �� ���� "�������"
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
    
    Range("W" & off).Value = ActiveWorkbook.Sheets("���������").Range("D12").Value        ' ���� ����������
    Range("X" & off & ":AA" & off).Select
    Selection.FillDown
    
    Range("AB" & off).Value = ActiveWorkbook.Sheets("���������").Range("D9").Value        ' ������
    Range("AC" & off).Value = ActiveWorkbook.Sheets("���������").Range("D10").Value       ' ����������
    Range("AD" & off).Value = ActiveWorkbook.Sheets("���������").Range("D11").Value       ' �������
        
    Range("AE" & off).Select
    Selection.FillDown
    
    kvnNames_GetSheet("��������_����������").Select
    Range("D1").Value = ActiveWorkbook.Sheets("���������").Range("D12").Value             ' ����
    Range("C5").Value = ActiveWorkbook.Sheets("���������").Range("D8").Value              ' ID ������
    Range("D32").Value = ActiveWorkbook.Sheets("���������").Range("D9").Value             ' ������

End Sub
Sub �������������_Sub(FlagFromJSON As Boolean)
Attribute �������������_Sub.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim idx_Spr_UserName, cUserName, idx, numDO
    Dim idx_SheetName_Retail, idx_Retail_ID, rw As Integer, mes
    Dim r1, r2, CoinID
    
    If FlagFromJSON Then
        kvnNames_GetSheet("��������_�_�������").Select
        CoinID = ActiveWorkbook.Sheets("���������").Range("D8").Value
        kvnNames_FindValue("�_�������!���������������", CoinID).Select
        rw = ActiveCell.row
    Else
        rw = ActiveCell.row
    End If
    
#If KIND_OF_INSTANCE = "DO" Then
    If ThisWorkbook.Name = ActiveWorkbook.Name Then
        MsgBox ("������ ����������� �� �������� ����� �� �����")
        Exit Sub
    End If
#End If
    
    idx_SheetName_Retail = kvnNames_GetProp("��������_�_�������")
    If ActiveSheet.Name <> idx_SheetName_Retail Then
        MsgBox ("������ ������������ ���������� �� ��������� ������� ����������� � ����� '" & idx_SheetName_Retail & "'!")
        Exit Sub
    End If
    
    idx_Spr_UserName = kvnNames_GetProp("idx_Spr_UserName")
    cUserName = GetUserName()
    If IsError(Application.Match(cUserName, Range(idx_Spr_UserName), 0)) = True Then
        MsgBox ("� ���������, � ��� ��� ���� �� ������������ ����������.")
        Exit Sub
    End If
    
    idx = Application.Match(cUserName, Range(idx_Spr_UserName), 0)
    numDO = Application.WorksheetFunction.Index(Range(kvnNames_GetProp("idx_Spr_DO_for_UserName")), idx, 0)
    
    If kvnNames_GetCell("�_�������!������", rw).Value <> kvnNames_GetProp("������_��������") Then
        MsgBox ("��� ������� ����� ������ �� ������ � ������� �� �������� '� �������'!")
        Exit Sub
    End If
    
    If kvnNames_GetCell("�_�������!�������������", rw).Value = kvnNames_GetProp("���������������������") Then
        MsgBox ("������� ������ �� ��������� ����������!")
        Exit Sub
    End If
    
    If numDO <> "���" Then
        If kvnNames_GetCell("�_�������!�������������", rw).Value <> numDO Then
            MsgBox ("��� ��������� ������ ������� ����� ������������� " & numDO & "!")
            Exit Sub
        End If
    End If
    
    ActiveWorkbook.Sheets("���������").Range("D3").Value = ActiveCell.row
    ActiveWorkbook.Sheets("���������").Range("D14").Value = idx + 1
    
    If Not FlagFromJSON Then
        AktSale.Show
        If Not FormResult Then Exit Sub
    End If
    
#If KIND_OF_INSTANCE = "DO" Then
    mes = CreateJSON(rw)
    MsgBox ("������� ���������." & vbCrLf & mes)
#Else
    Call CopyToSale(rw)
    Call Send_Email_Using_Keys
    Sleep 200
    'MsgBox ("������� ���������")
#End If

End Sub
Private Sub Send_Email_Using_Keys()
    Dim Mail_Object As String
    Dim Email_Subject, Email_Send_To, Email_Cc, Email_Bcc, Email_Body As String
    Dim hwnd As Long, nm
    
    Email_Subject = "������������ (���������� ������ �" & Range("C5").Value & ")"
    Email_Send_To = kvnNames_GetProp("������������������������")
        
    'Email_Cc = "exceltipmail@gmail.com "
    'Email_Bcc = "exceltipmail@gmail.com "
    'Email_Body = "%0D%0A" & Sheets("���������").Range("D15").Value & "%0D%0A"
    Mail_Object = "mailto:" & Email_Send_To & "?subject=" & Email_Subject '& " &body=" & Email_Body

    ' *****************************************
    Sheets("���� ����������").Select
' Sheets("������������").Copy
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
    '    '   Application.SendKeys "%(�)"
    '    '   Application.SendKeys ("{ENTER}")
    '    Application.StatusBar = False
    'Else
    '    Application.StatusBar = hwnd & " -> " & nm
    'End If
    
'debugs:
'    If Err.Description <> "" Then MsgBox Err.Description
End Sub
' ****************************************************************************************************************
' C�������� ��������� (�����������_�����, ���������������_�_���������������, ��������������������������������, ����������������) �������� ������ � ��������� Main
' ****************************************************************************************************************
#If KIND_OF_INSTANCE = "MAIN" Then
Sub ����������_�����()
    Dim fName, CurName, vbFileName
    
    fName = kvnNames_GetProp("�������������������������") & kvnNames_GetProp("�����������������������")
    CurName = ActiveWorkbook.Name
    
    If Dir(fName) <> "" Then                            ' ���� ����������
        SetAttr fName, vbNormal                         ' ������ ������� Read Only
        Kill (fName)                                    ' ������ ����
    End If

    Windows(ThisWorkbook.Name).Activate                 ' ��������� ������ � ������� ����� � �����
    Sheets(Array("���������� �����", "� �������", "����������� ����", "���������", "kvnNames")).Copy
    kvnNames_GetSheet("��������_�_�������").Columns("S:X").EntireColumn.Hidden = True ' ������ ������� � S �� X
    Range("A1").Select
    
    ' �������� ������ �� ���������� �����, ��� ��������
    kvnNames_GetSheet("��������_���������������").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' ��������� ����
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=fName, FileFormat:=xlExcel8, CreateBackup:=False 'xlOpenXMLWorkbook
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    SetAttr fName, vbReadOnly
    
    Windows(CurName).Activate
    kvnNames_GetSheet("��������_���������������").Select
    
    MsgBox ("����� " & fName & " ���������")
End Sub

' *.JSON
' -----------------------------------------
'{
'    "ID_������": "1347",
'    "������": "������� �.�.",
'    "������": "������ ���� ��������",
'    "�������": "2-33-223",
'    "��������������": "30.05.2017",
'    "Login": "KozulinVN",
'    "User": "50"
'}
' -----------------------------------------
Sub ���������������_�_���������������()
    Dim fName, fPath, t
    Dim cName, cValue
    Dim PathChange
    
    PathChange = kvnNames_GetProp("����������JSON")
    
    fPath = PathChange & "S_*.JSON"
    fName = Dir(fPath, vbNormal)
    If (fName = "") Then
        MsgBox ("������������� ���. �����, �� ����� ������ �� �������.")
        Exit Sub
    Else
        MsgBox ("��������� ���� " & fName)
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
            t = Application.WorksheetFunction.Match(cName, ActiveWorkbook.Sheets("���������").Range("A8:A14"), 0)
            If ActiveWorkbook.Sheets("���������").Range("D7").Offset(t, 0).NumberFormat = "m/d/yyyy" Then
                cValue = CDate(cValue)
            End If
            ActiveWorkbook.Sheets("���������").Range("D7").Offset(t, 0).Value = cValue
        End If
    Wend
    Close #1
    
    Call �������������_Sub(True)
    Kill PathChange & fName
    
    If Dir() <> "" Then
        MsgBox ("� �������� ��� ���� ����� ��� ���������. ��������� ������ ��������.")
    Else
        MsgBox ("��� ����� � �������� ����������.")
    End If
End Sub
Sub ��������������������������������()
    Dim id, rw As Integer
    
    If ActiveSheet.Name <> kvnNames_GetProp("���������������") Then
        MsgBox ("����������� ������ �� ����� '" & kvnNames_GetProp("���������������") & "'")
        Exit Sub
    End If
    
    If kvnNames_GetCell("�������!������").Value = "��������" Then
        MsgBox ("������ ��� �������� � �����")
        Exit Sub
    Else
        kvnNames_GetCell("�������!������").Value = "��������"
    End If
    
    id = kvnNames_GetCell("�������!���������������").Value
    rw = kvnNames_FindValue("�_�������!���������������", id).row
    
    kvnNames_GetSheet("��������_�_�������").Activate
    Rows(rw & ":" & rw).Select
    If kvnNames_GetCell("�_�������!������", rw).Value <> "�������" Then
        MsgBox ("�� �������� ������� �� ��������� ������!")
        Exit Sub
    End If
        
    If MsgBox("������ ������ " & t, vbYesNo + vbDefaultButton2) = vbYes Then
        Selection.Delete Shift:=xlUp
    End If
    kvnNames_GetSheet("���������������").Activate
    
    MsgBox ("������ " & id & " ������� �� ����� '� �������'")
    
End Sub
Sub ����������������()
    Dim mName
    
    mName = kvnNames_GetProp("������������������������")
    
    If ActiveWorkbook.Name <> mName Then
        MsgBox ("����������� ������ �� '" & mName & "'")
        Exit Sub
    End If
    
    ' ������ ��� ������ �������
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
    ' ��������� ���� �������
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
    ActiveWorkbook.VBProject.VBComponents("���������").CodeModule.ReplaceLine 9, "'#Const KIND_OF_INSTANCE = ""MAIN"""
    ActiveWorkbook.VBProject.VBComponents("���������").CodeModule.ReplaceLine 10, "#Const KIND_OF_INSTANCE = ""DO"""
    MsgBox ("������� �����������")
End Sub

#End If
' ****************************************************************************************************************
' ���������� ��������� (�����������_�����, ���������������_�_���������������, ��������������������������������, ����������������) �������� ������ � ��������� Main
' ****************************************************************************************************************

Sub �������������()
#If KIND_OF_INSTANCE = "DO" Then
    Dim i As Integer, flag As Boolean
    Dim �����������������������
    
    ����������������������� = kvnNames_GetProp("�����������������������")
    If ThisWorkbook.Name = ActiveWorkbook.Name Then                     ' ������ ������� �������������� � ����� ��������
        i = 1                                                           ' ���������� ����� ����� ���� ������
        flag = False
        While (i <= Application.Workbooks.Count) And (flag = False)
            If Application.Workbooks(i).Name = ����������������������� Then
                Windows(�����������������������).Activate
                kvnNames_GetSheet("�����������������������������������").Select
                flag = True
            Else
                i = i + 1
            End If
        Wend
        If Not flag Then
            Workbooks.Open Filename:=kvnNames_GetProp("�������������������������") & �����������������������, UpdateLinks:=0
            kvnNames_GetSheet("�����������������������������������").Select
        End If
    End If
#End If
    Call �������������_Sub(False)
    
End Sub
