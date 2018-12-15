Attribute VB_Name = "Module2"
Function transalte_using_vba(str) As String
' Tools Refrence Select Microsoft internet Control


    Dim IE As Object, i As Long
    Dim inputstring As String, outputstring As String, text_to_convert As String, result_data As String, CLEAN_DATA

    Set IE = CreateObject("InternetExplorer.application")
    '   TO CHOOSE INPUT LANGUAGE

    inputstring = "auto"

    '   TO CHOOSE OUTPUT LANGUAGE

    outputstring = "en"

    text_to_convert = str

    'open website

    IE.Visible = False
    IE.Navigate "http://translate.google.com/#" & inputstring & "/" & outputstring & "/" & text_to_convert

    Do Until IE.readyState = 4
        DoEvents
    Loop

    Application.Wait (Now + TimeValue("0:00:5"))

    Do Until IE.readyState = 4
        DoEvents
    Loop
    'MsgBox (IE.Document)
    

    CLEAN_DATA = Split(Application.WorksheetFunction.Substitute(IE.Document.getElementById("result_box").innerHTML, "</SPAN>", ""), "<")

    For j = LBound(CLEAN_DATA) To UBound(CLEAN_DATA)
        result_data = result_data & Right(CLEAN_DATA(j), Len(CLEAN_DATA(j)) - InStr(CLEAN_DATA(j), ">"))
    Next


    IE.Quit
    transalte_using_vba = result_data


End Function
