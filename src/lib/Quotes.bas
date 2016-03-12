Attribute VB_Name = "Quotes"
Option Explicit

'Reference: http://social.msdn.microsoft.com/Forums/en-US/isvvba/thread/bd0ee306-7bb5-4ce4-8341-edd9475f84ad
Dim downloadPath As String
'downloadPath = Sheets("手動下載").Range("g1") & "\"  'Application.ActiveWorkbook.path & "\import\"

Function saveCSVbyURL(myURL As String, outFile As String) As Boolean
    Dim oStream As Object           'ADODB.Stream
    Dim WinHttpReq As Object            'XMLHTTP
    
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
        
        On Error Resume Next
        Kill outFile
        On Error GoTo 0
        
        .SaveToFile outFile
        .Close
        
        saveCSVbyURL = verifyFile(outFile)
        
    End With
    Set WinHttpReq = Nothing
    Set oStream = Nothing

End Function

'-測試是否下載正常: 去除表頭大小
Function verifyFile(fileIdx As String, Optional fileSize As Long = 50000)
     '測試檔案大小
    Dim fs, f, s

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(fileIdx)
    s = UCase(f.Size)
    verifyFile = True
    
    If s < fileSize Then      '檔案大小應為50k以上
      verifyFile = False
      's = "下載上市/櫃" & UCase(f.Name) & "盤後資料有誤，確認日期或稍後再下載。"
      'MsgBox s, 0, "盤後資料錯誤"
      fs.DeleteFile (fileIdx)     '移除錯誤的檔案
        
    End If
End Function

Function downloadQuotesFromTSE(selDate As String)
    Dim myURL As String
    Dim fileIdx As String, outPath As String
    
    myURL = "http://www.twse.com.tw/ch/trading/exchange/MI_INDEX/MI_INDEX3_print.php?" & _
                    "genpage=genpage/Report" & VBA.Left(selDate, 6) & "/A112" & selDate & "ALL_1.php&type=csv"
       
    outPath = Application.ActiveWorkbook.path & "\import\"   ' Sheets("手動下載").Range("g1") & "\"
    fileIdx = outPath & "A112" & selDate & "ALL_1.csv"
       
    downloadQuotesFromTSE = saveCSVbyURL(myURL, fileIdx)

End Function

Function downloadQuotesFromOTC(selDate As String)
    Dim st, sd
    Dim myURL As String
    Dim fileIdx As String, outPath As String
    
    sd = VBA.Left(selDate, 4) - 1911 & "/" & VBA.Mid(selDate, 5, 2) & "/" & VBA.Right(selDate, 2)
    st = VBA.Left(selDate, 4) - 1911 & VBA.Right(selDate, 4)
    'myURL = "http://www.gretai.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/RSTA3104_" & st & ".csv"
    'myURL = "http://www.gretai.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/stk_quote_download.php?d=" & sd & "&s=0,asc,0"
    myURL = "http://www.tpex.org.tw/ch/stock/aftertrading/DAILY_CLOSE_quotes/stk_quote_download.php?d=" & sd & "&s=0,asc,0"
    
    outPath = Application.ActiveWorkbook.path & "\import\"   ' Sheets("手動下載").Range("g1") & "\"
    fileIdx = outPath & "RSTA3104_" & st & ".csv"
       
    downloadQuotesFromOTC = saveCSVbyURL(myURL, fileIdx)

End Function


