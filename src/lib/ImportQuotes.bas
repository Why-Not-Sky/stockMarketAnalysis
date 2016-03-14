Attribute VB_Name = "mm_transport"
Option Compare Database

Function export250()
'--- 把250天以前的資料存到 export250
Dim begDate As String
Dim sql As String
Dim fileName As String

begDate = getPrevStockDate(Date, 250)

fileName = "export_" & chiDate(begDate)
sql = "select * into " & fileName & " from stk where dte < #" & begDate & "#"
Debug.Print sql

DoCmd.RunSQL (sql)
MsgBox "已將" & begDate & "之前資料存到 " & fileName & " 請先檢查過後再執行[清除250日前資料]動作"
End Function
Function delete250()
'--- 把250天以前的資料刪除
Dim begDate As String
Dim sql As String
Dim fileName As String

begDate = getPrevStockDate(Date, 250)
sql = "delete from stk where dte < #" & begDate & "#"
Debug.Print sql

DoCmd.RunSQL (sql)
MsgBox "已將" & begDate & "之前資料刪除，請執行[工具][資料庫公用程式][壓縮及修復資料庫]讓檔案變小 "
End Function

Function parseRSTdate(s1)
'line=2 資料日期:102/12/23
 ix = InStr(s1, ":")
  
 yy = mid(s1, ix + 1, 3)
 mmdd = mid(s1, ix + 4)
 sDate = (1911 + yy) & mmdd
 parseRSTdate = sDate
 Debug.Print sDate
End Function
    
Function RSTAimp1(source)
' --- 匯入一個上櫃csv檔到主檔stk
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
    If lines = 0 Then   'line=1 102年11月22日 櫃檯買賣中心證券行情
       ix = InStr(s1, "年")

        yy = mid(s1, 1, ix - 1)
        mm = mid(s1, InStr(s1, "年") + 1, 2)
        dd = mid(s1, InStr(s1, "月") + 1, 2)
        sDate = (1911 + yy) & "/" & mm & "/" & dd
        Debug.Print sDate
    End If
    
    'new formate
    If lines = 2 Then sDate = parseRSTdate(s1)

    If InStr(s1, "上櫃家數") <> 0 Then
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
            '代號0,名稱1,收盤2 ,漲跌3,開盤4 ,最高5 ,最低6,均價7 ,成交股數8
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
                '--- 同時更新股票代號檔
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
'--- 將dirImport內的所有 RSTA*.csv上櫃檔匯入到主檔 stk，同時將該檔移到 dircomplete
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
    MsgBox "請將欲轉入之xls 放在" & dirImport
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
Debug.Print "Done!RSTA*.csv Import" & vbCrLf & " 檔案已經搬移到" & dirComplete

End Function


Function A11imp1(source)
' --- 匯入一個上市csv檔到主檔stk
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
    If Not isBegin And InStr(s1, "每日收盤行情") <> 0 Then
        isBegin = True
        s1 = Replace(s1, """", "")
        ix = InStr(s1, "年")
        
        yy = mid(s1, 1, ix - 1)
        
        mm = mid(s1, InStr(s1, "年") + 1, 2)
        dd = mid(s1, InStr(s1, "月") + 1, 2)
        sDate = (1911 + yy) & "/" & mm & "/" & dd
        
        
        DoCmd.RunSQL ("delete from stk where dte=#" & sDate & "#")
        
        Debug.Print sDate
        inpFi.SkipLine        ' 標題欄
    ElseIf isBegin Then
           
        '--- 解決字串中的,號，如"200,450,000"==> 200450000
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
            
            '*** 判斷該行是否為所需要的？剔除權證以及非資料行
            isValid = True
            If UBound(aa) < 8 Then
                isValid = False
            End If
            If Left(aa(0), 1) = "0" And Len(aa(0)) <> 4 Then
                isValid = False
            End If
            
            
            If isValid Then
            
            '證券代號0,證券名稱1,成交股數2,成交筆數3,成交金額4,開盤價5,最高價6,最低價7,收盤價8
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
                '--- 同時更新股票代號檔
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
'--- 將dirImport內的所有 A11*.csv上市檔匯入到主檔 stk，同時將該檔移到 dircomplete
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
    MsgBox "請將欲轉入之xls 放在" & dirImport
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
Debug.Print "Done!A11*.csv Import" & vbCrLf & " 檔案已經搬移到" & dirComplete

End Function
Function importDaily()
    Call A11Import
    Call RSTAImport
    MsgBox "Done!"
End Function
