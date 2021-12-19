Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub getSlumpedStocks()
    If MsgBox("点击确认将开始抓取全部A股数据，时间较长（请看左下角进度条，抓取中请勿操作）", vbYesNo) <> vbYes Then Exit Sub
    
    iOriginIndex = 2    '开始index
    iIndex = iOriginIndex
    iSelfIndex = 2
    iMaxRow = Sheet9.[a65536].End(xlUp).Row + iIndex
    
    'clear
    iMaxLine = Sheet11.Cells.SpecialCells(xlCellTypeLastCell).Column
    Sheet11.Range(Cells(iOriginIndex, 1), Cells(iMaxRow, iMaxLine)).ClearContents

    ' 日期
    Dim arrYear As Variant
    arrYear = Array(2017, 2018, 2019, 2020, 2021)
        
    Dim obj
    Set obj = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim dKlineDate As Date
    Dim sJsonStr, sDate As String
    Dim iPrice, iHighestPriceInYear As Single

    Do
        Sheet11.Cells(iSelfIndex, "a") = Sheet9.Cells(iIndex, 1)
        Sheet11.Cells(iSelfIndex, "b") = Sheet9.Cells(iIndex, 2)
        
        code = Sheet9.Cells(iIndex, 1)
        marketCode = getMarketCode(code)
                
        ' 周K
        ' Url = "http://74.push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112407697307125612478_1638805346501&secid=" & marketCode & "." & code & "&ut=fa5fd1943c7b386f172d6893dbfba10b&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5%2Cf6&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58%2Cf59%2Cf60%2Cf61&klt=101&fqt=0&beg=20170101&end=20500101&smplmt=460&lmt=1000000&_=1638805346528"
        
        ' 月k
        Url = "http://90.push2his.eastmoney.com/api/qt/stock/kline/get?cb=jQuery112402816109881259974_1638893884237&secid=" & marketCode & "." & code & "&ut=fa5fd1943c7b386f172d6893dbfba10b&fields1=f1%2Cf2%2Cf3%2Cf4%2Cf5%2Cf6&fields2=f51%2Cf52%2Cf53%2Cf54%2Cf55%2Cf56%2Cf57%2Cf58%2Cf59%2Cf60%2Cf61&klt=103&fqt=0&beg=20170101&end=20500101&smplmt=460&lmt=1000000&_=1638893884280"
        
        obj.Open "GET", Url, True
        obj.send
        obj.WaitForResponse
        t1 = BytesToBstr(obj.ResponseBody, "UTF-8")
        t1 = Split(t1, "(")(1)
        Set x = CreateObject("ScriptControl"): x.Language = "JScript"
        x.AddCode ("var query = (" & t1)

        '开始处理K线数据
        rData = x.Eval("query.data")
        If rData <> "Null" Then
            ikLineLen = x.Eval("query.data.klines.length")
            For i = 0 To UBound(arrYear)
                iYear = arrYear(i)

                iHighestPriceInYear = 0
                For j = 0 To ikLineLen - 1
                    ' 时间，开盘价，收盘价，最高价，最低价
                    sJsonStr = x.Eval("query.data.klines[" & j & "]")
                    sDate = Split(sJsonStr, ",")(0)
                    dKlineDate = DateValue(sDate)

                    if Year(dKlineDate) = iYear Then
                        iPrice = Val(Split(sJsonStr, ",")(3))
                        if iPrice > iHighestPriceInYear Then
                            iHighestPriceInYear = iPrice
                        end if
                    end if
                Next j

                Sheet11.Cells(iIndex, 3 + i) = iHighestPriceInYear
            Next i
        end if
        Application.StatusBar = GetProgress(iIndex - iOriginIndex, iMaxRow - iOriginIndex)
        iIndex = iIndex + 1      
        iSelfIndex = iSelfIndex + 1 
        Sleep 10
    Loop Until iIndex > iMaxRow
    MsgBox "A股所有数据抓取完毕", , "提示"
End Sub
