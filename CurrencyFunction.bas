Attribute VB_Name = "CurrencyFunction"
Public Function wConvertCurrency(currency1, currency2)
    Dim yahooHTTP As New WinHttp.WinHttpRequest
    yahooHTTP.Open "GET", "http://quote.yahoo.com/d/quotes.csv?s=" & currency1 & currency2 & "=X&f=l1"
    yahooHTTP.Send
    wConvertCurrency = CDbl(yahooHTTP.responseText)
    'MsgBox yahooHTTP.ResponseText
End Function


