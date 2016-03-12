Attribute VB_Name = "Quotes"
Option Explicit

'Reference: http://social.msdn.microsoft.com/Forums/en-US/isvvba/thread/bd0ee306-7bb5-4ce4-8341-edd9475f84ad

Function saveCSVbyURL(myURL As String, tradeDate As String, outFile As String) As Boolean

    'Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    Set WinHttpReq = CreateObject("MSXML2.XMLHTTP")
    With WinHttpReq
        .Open "GET", myURL, False
        .Send
        myURL = .responseText
    End With
    
    Set oStream = CreateObject("ADODB.Stream")
    With oStream
        .Open
        .Type = 1
        .Write WinHttpReq.responseBody
        'fileIdx = Sheets("手動下載").Range("f2") & "\A112" & fileIdx & "ALL_1.csv"
        On Error Resume Next
        Kill fileIdx
        On Error GoTo 0
        .SaveToFile outFile
        .Close
        
        verifyFile fileIdx
        
    End With
    Set WinHttpReq = Nothing
    Set oStream = Nothing

End Function

'-測試是否下載正常
Function verifyFile(fileIdx As String)
     '測試檔案大小
    Dim fs, f, s

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(fileIdx)
    s = UCase(f.Size)
    verifyFile = True
    
    If s < 500000 Then      '檔案大小500k以上
      verifyFile = False
      s = "下載上市/櫃" & UCase(f.Name) & "盤後資料有誤，確認日期或稍後再下載。"
      MsgBox s, 0, "盤後資料錯誤"
        fs.DeleteFile (fileIdx)     '移除錯誤的檔案
        
    End If
End Function


Function saveCSVfmURL(selDate As String)
Dim st, sd
Dim myURL As String
Dim oStream As Object           'ADODB.Stream
Dim WinHttpReq As Object            'XMLHTTP
Dim fileIdx As String
fileIdx = selDate


myURL = "http://www.twse.com.tw/ch/trading/exchange/MI_INDEX/MI_INDEX3_print.php?" & _
                "genpage=genpage/Report" & VBA.Left(fileIdx, 6) & "/A112" & fileIdx & "ALL_1.php&type=csv"


 
'Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
Set WinHttpReq = CreateObject("MSXML2.XMLHTTP")
With WinHttpReq
    .Open "GET", myURL, False
    .Send
    myURL = .responseText
End With
Set oStream = CreateObject("ADODB.Stream")
With oStream
    .Open
    .Type = 1
    .Write WinHttpReq.responseBody
    fileIdx = Sheets("手動下載").Range("g1") & "\A112" & fileIdx & "ALL_1.csv"
    On Error Resume Next
    Kill fileIdx
    On Error GoTo 0
    .SaveToFile fileIdx
    .Close
    
    verifyFile fileIdx
End With
Set WinHttpReq = Nothing
Set oStream = Nothing


fileIdx = selDate
sd = VBA.Left(fileIdx, 4) - 1911 & "/" & VBA.Mid(fileIdx, 5, 2) & "/" & VBA.Right(fileIdx, 2)
st = VBA.Left(fileIdx, 4) - 1911 & VBA.Right(fileIdx, 4)
'myURL = "http://www.gretai.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/RSTA3104_" & st & ".csv"
'myURL = "http://www.gretai.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/stk_quote_download.php?d=" & sd & "&s=0,asc,0"
myURL = "http://www.tpex.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/stk_quote_download.php?d=" & sd & "&s=0,asc,0"
'http://www.gretai.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/stk_quote_download.php?d=103/01/15&s=0,asc,0

'Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
Set WinHttpReq = CreateObject("MSXML2.XMLHTTP")
With WinHttpReq
    .Open "GET", myURL, False
    .Send
    myURL = .responseText
End With

Set oStream = CreateObject("ADODB.Stream")
With oStream
    .Open
    .Type = 1
    .Write WinHttpReq.responseBody
    fileIdx = Sheets("手動下載").Range("g1") & "\RSTA3104_" & st & ".csv"
    On Error Resume Next
    Kill fileIdx
    
    On Error GoTo 0
    .SaveToFile fileIdx
    .Close
    
    saveCSVfmURL = verifyFile(fileIdx)
    
End With
Set WinHttpReq = Nothing
Set oStream = Nothing


End Function

