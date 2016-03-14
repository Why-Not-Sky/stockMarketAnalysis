Attribute VB_Name = "mm_transport"
Option Compare Database

Function export250()
'--- ��250�ѥH�e����Ʀs�� export250
Dim begDate As String
Dim sql As String
Dim fileName As String

begDate = getPrevStockDate(Date, 250)

fileName = "export_" & chiDate(begDate)
sql = "select * into " & fileName & " from stk where dte < #" & begDate & "#"
Debug.Print sql

DoCmd.RunSQL (sql)
MsgBox "�w�N" & begDate & "���e��Ʀs�� " & fileName & " �Х��ˬd�L��A����[�M��250��e���]�ʧ@"
End Function
Function delete250()
'--- ��250�ѥH�e����ƧR��
Dim begDate As String
Dim sql As String
Dim fileName As String

begDate = getPrevStockDate(Date, 250)
sql = "delete from stk where dte < #" & begDate & "#"
Debug.Print sql

DoCmd.RunSQL (sql)
MsgBox "�w�N" & begDate & "���e��ƧR���A�а���[�u��][��Ʈw���ε{��][���Y�έ״_��Ʈw]���ɮ��ܤp "
End Function

Function parseRSTdate(s1)
'line=2 ��Ƥ��:102/12/23
 ix = InStr(s1, ":")
  
 yy = mid(s1, ix + 1, 3)
 mmdd = mid(s1, ix + 4)
 sDate = (1911 + yy) & mmdd
 parseRSTdate = sDate
 Debug.Print sDate
End Function
    
Function RSTAimp1(source)
' --- �פJ�@�ӤW�dcsv�ɨ�D��stk
Dim fs2 As New Scripting.FileSystemObject
Dim RsStk As DAO.Recordset
Dim RsStkId As DAO.Recordset
Dim inpFi As TextStream
Dim lines As Integer
Dim s1 As String
Dim yy As String, mm As String, dd As String
Dim sDate As String
Dim aa
Dim i As Integer, ix As Integer
Dim s2 As String
Dim isdel As Boolean
Dim ch

Set inpFi = fs2.OpenTextFile(source)
Set RsStk = CurrentDb.OpenRecordset("select * from stk")
lines = 0
Do While Not inpFi.AtEndOfStream
    lines = lines + 1
    s1 = inpFi.ReadLine
    If lines = 0 Then   'line=1 102�~11��22�� �d�i�R�椤���Ҩ�污
       ix = InStr(s1, "�~")

        yy = mid(s1, 1, ix - 1)
        mm = mid(s1, InStr(s1, "�~") + 1, 2)
        dd = mid(s1, InStr(s1, "��") + 1, 2)
        sDate = (1911 + yy) & "/" & mm & "/" & dd
        Debug.Print sDate
    End If
    
    'new formate
    If lines = 2 Then sDate = parseRSTdate(s1)

    If InStr(s1, "�W�d�a��") <> 0 Then
        Exit Do
    End If
    If lines >= 4 Then
            '**************************************
            isdeli = False
            s2 = ""
            For i = 1 To Len(s1)
                ch = mid(s1, i, 1)
                If mid(s1, i, 1) = """" Then
                    isdeli = Not isdeli
                Else
                    If isdeli Then
                        If ch <> "," Then
                            s2 = s2 & ch
                        End If
                    Else
                        s2 = s2 & ch
                    End If
                End If
            Next
            s2 = Replace(s2, "--", "00")
            'Debug.Print s2
            '*************************************
    
        aa = Split(s2, ",")
        If UBound(aa) < 9 Then
            Debug.Print s1
        ElseIf Len(aa(0)) = 4 Then
            '�N��0,�W��1,���L2 ,���^3,�}�L4 ,�̰�5 ,�̧C6,����7 ,����Ѽ�8
            RsStk.AddNew
            RsStk("dte") = CDate(sDate)
            RsStk("stockid") = aa(0)
            RsStk("price") = aa(2)
            RsStk("p_open") = aa(4)
            RsStk("p_high") = aa(5)
            RsStk("p_low") = aa(6)
            s2 = Replace(aa(8), ",", "")
            s2 = Replace(s2, """", "")
            RsStk("vol") = s2 / 1000
            RsStk.Update
                '--- �P�ɧ�s�Ѳ��N����
                Set RsStkId = CurrentDb.OpenRecordset("select * from stkid where stockid='" & aa(0) & "'")
                If RsStkId.EOF Then
                    DoCmd.RunSQL ("insert into stkid (stockid,stkname,CL) values('" & aa(0) & "','" & aa(1) & "','2')")
                End If
            
        End If
    End If
Loop
RsStk.Close
Debug.Print source
End Function
Function RSTAImport()
'--- �NdirImport�����Ҧ� RSTA*.csv�W�d�ɶפJ��D�� stk�A�P�ɱN���ɲ��� dircomplete
Dim fs As New Scripting.FileSystemObject
Dim outFi As TextStream, errFi As TextStream

Dim dirImport As String, dirComplete As String
Dim fi As File
Dim strRoot As String
Dim fo_root As Folder
Dim fo_Dir As Folder
Dim ext As String

Dim source As String, Target As String, stockid As String

strRoot = CurrentProject.path
dirImport = strRoot & "\import"
dirComplete = strRoot & "\complete"

Set outFi = fs.CreateTextFile(strRoot & "\" & "log.txt", True)
Set errFi = fs.CreateTextFile(strRoot & "\" & "errlog.txt", True)

Set fo_root = fs.GetFolder(strRoot)

If Not fs.FolderExists(dirImport) Then
    MsgBox "�бN����J��xls ��b" & dirImport
    Exit Function
End If

If Not fs.FolderExists(dirComplete) Then
    fs.CreateFolder (dirComplete)
End If

Set fo_Dir = fs.GetFolder(dirImport)

For Each fi In fo_Dir.Files
    ext = mid(fi.Name, InStr(fi.Name, "."), 4)
    If Left(fi.Name, 4) = "RSTA" And LCase(ext) = ".csv" Then
        Debug.Print fi.Name
        source = fi.path
        Target = dirComplete & "\" & fi.Name
        fs.CopyFile source, Target, True
        Call RSTAimp1(source)
        fs.DeleteFile source
    End If
Next
outFi.Close
errFi.Close
Debug.Print "Done!RSTA*.csv Import" & vbCrLf & " �ɮפw�g�h����" & dirComplete

End Function


Function A11imp1(source)
' --- �פJ�@�ӤW��csv�ɨ�D��stk
Dim fs2 As New Scripting.FileSystemObject
Dim RsStk As DAO.Recordset
Dim RsStkId As DAO.Recordset

Dim inpFi As TextStream
Dim s1 As String
Dim yy As String, mm As String, dd As String
Dim ix As Integer
Dim sDate As String
Dim aa
Dim i As Integer
Dim s2 As String
Dim isBegin  As Boolean
Dim isdelim As Boolean
Dim lines, ch
Dim isValid As Boolean



Set inpFi = fs2.OpenTextFile(source)
Set RsStk = CurrentDb.OpenRecordset("select * from stk")
isBegin = False
lines = 0

Do While Not inpFi.AtEndOfStream
    s1 = inpFi.ReadLine
    If Not isBegin And InStr(s1, "�C�馬�L�污") <> 0 Then
        isBegin = True
        s1 = Replace(s1, """", "")
        ix = InStr(s1, "�~")
        
        yy = mid(s1, 1, ix - 1)
        
        mm = mid(s1, InStr(s1, "�~") + 1, 2)
        dd = mid(s1, InStr(s1, "��") + 1, 2)
        sDate = (1911 + yy) & "/" & mm & "/" & dd
        
        
        DoCmd.RunSQL ("delete from stk where dte=#" & sDate & "#")
        
        Debug.Print sDate
        inpFi.SkipLine        ' ���D��
    ElseIf isBegin Then
           
        '--- �ѨM�r�ꤤ��,���A�p"200,450,000"==> 200450000
        'If Mid(s1, 1, 1) <> "0" Then
            lines = lines + 1
            'If lines > 5 Then
            '    Exit Do
            'End If
            '**************************************
            isdeli = False
            s2 = ""
            For i = 1 To Len(s1)
                ch = mid(s1, i, 1)
                If mid(s1, i, 1) = """" Then
                    isdeli = Not isdeli
                Else
                    If isdeli Then
                        If ch <> "," Then
                            s2 = s2 & ch
                        End If
                    Else
                        s2 = s2 & ch
                    End If
                End If
            Next
            s2 = Replace(s2, "--", "00")
            s2 = Replace(s2, "=", "")
            'Debug.Print s2
            '*************************************
            aa = Split(s2, ",")
            
            '*** �P�_�Ӧ�O�_���һݭn���H�簣�v�ҥH�ΫD��Ʀ�
            isValid = True
            If UBound(aa) < 8 Then
                isValid = False
            End If
            If Left(aa(0), 1) = "0" And Len(aa(0)) <> 4 Then
                isValid = False
            End If
            
            
            If isValid Then
            
            '�Ҩ�N��0,�Ҩ�W��1,����Ѽ�2,���浧��3,������B4,�}�L��5,�̰���6,�̧C��7,���L��8
                RsStk.AddNew
                RsStk("dte") = CDate(sDate)
                RsStk("stockid") = aa(0)
                RsStk("price") = "0" & aa(8)
                RsStk("p_open") = aa(5)
                RsStk("p_high") = aa(6)
                RsStk("p_low") = aa(7)
                s2 = Replace(aa(2), ",", "")
                s2 = Replace(s2, """", "")
                s2 = Replace(s2, " ", "")
                RsStk("vol") = s2 / 1000
                RsStk.Update
                '--- �P�ɧ�s�Ѳ��N����
                Set RsStkId = CurrentDb.OpenRecordset("select * from stkid where stockid='" & aa(0) & "'")
                If RsStkId.EOF Then
                    DoCmd.RunSQL ("insert into stkid (stockid,stkname,CL) values('" & aa(0) & "','" & aa(1) & "','1')")
                End If
            End If
        'End If ' Mid(s1, 1, 1) = "0" T
    End If 'isBegin
Loop
RsStk.Close
Debug.Print source
End Function
Function A11Import()
'--- �NdirImport�����Ҧ� A11*.csv�W���ɶפJ��D�� stk�A�P�ɱN���ɲ��� dircomplete
Dim fs As New Scripting.FileSystemObject
Dim outFi As TextStream, errFi As TextStream

Dim dirImport As String, dirComplete As String
Dim fi As File
Dim strRoot As String
Dim fo_root As Folder
Dim fo_Dir As Folder
Dim ext As String

Dim source As String, Target As String, stockid As String

strRoot = CurrentProject.path
dirImport = strRoot & "\import"
dirComplete = strRoot & "\complete"

Set outFi = fs.CreateTextFile(strRoot & "\" & "log.txt", True)
Set errFi = fs.CreateTextFile(strRoot & "\" & "errlog.txt", True)

Set fo_root = fs.GetFolder(strRoot)

If Not fs.FolderExists(dirImport) Then
    MsgBox "�бN����J��xls ��b" & dirImport
    Exit Function
End If

If Not fs.FolderExists(dirComplete) Then
    fs.CreateFolder (dirComplete)
End If

Set fo_Dir = fs.GetFolder(dirImport)

For Each fi In fo_Dir.Files
    ext = mid(fi.Name, InStr(fi.Name, "."), 4)
    If Left(fi.Name, 3) = "A11" And LCase(ext) = ".csv" Then
        Debug.Print fi.Name
        source = fi.path
        Target = dirComplete & "\" & fi.Name
        fs.CopyFile source, Target, True
        Call A11imp1(source)
        fs.DeleteFile source
    End If
Next
outFi.Close
errFi.Close
Debug.Print "Done!A11*.csv Import" & vbCrLf & " �ɮפw�g�h����" & dirComplete

End Function
Function importDaily()
    Call A11Import
    Call RSTAImport
    MsgBox "Done!"
End Function
