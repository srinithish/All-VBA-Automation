Attribute VB_Name = "Module1"
Function translate(fromStr)
'enable microsoft scripting runtime
'enable microsoft winHTTPServices 5.1

Dim Request  As New WinHttp.WinHttpRequest


Dim URL As String, langFrom As String, langTo As String, URL2 As String, JSON As Object, textToConv As String
langFrom = "auto"
langTo = "es"
textToConv = fromStr
'URL = "https://translate.google.com/#" & langFrom & "/" & langTo & "/" & fromStr
URL2 = "https://translation.googleapis.com/language/translate/v2?" & "&key=AIzaSyDnQ5spbz4YqySENebtugdZlLyOMMLAG9I" & "&source=" & "&target=en" & "&q=" & textToConv & "&format=text"


Request.Open "GET", URL2, False
Request.Send
responseString = Request.responseText
'Request.responseText
Set JSON = ParseJson(responseString)

'error handling displaays a msgbox with the error
If responseString Like "*error*" Then
MsgBox JSON("error")("errors")(1)("message")
Else
translate = JSON("data")("translations")(1)("translatedText")
End If


'Set JSON = ParseJson(Request.responseText)
'For Each Item In JSON
'ErrorString = JSON("error")("errors")(1)("message")
'
''If IsNothing Then
'MsgBox ErrorString
''End If
'MsgBox JSON("data")("translations")(1)("translatedText")

'responseString = JSON("data")("translations")(1)("translatedText")
'MsgBox JSON("error")("errors")(1)("message")

'error handling
'If responseString = "" Then
'MsgBox JSON("error")("errors")(1)("message")
'End If

'translate = responseString

End Function
            
    
