Attribute VB_Name = "controlWebsite"
Sub My()
Dim str As Variant
Dim IE As Object
Set IE = New InternetExplorer
IE.Visible = True
IE.Navigate "http://ktgis.net/gcode/geocoding.html"

Do While IE.Busy: DoEvents: Loop

With IE.Document

    Set someThing = .getElementsByTagName("textarea")

    For Each t In someThing 'to entere the adrresses in the text box
        If (t.getAttribute("name") = "address") Then
           Do While IE.Busy: DoEvents: Loop

           t.Value = "df" & vbCrLf & "something" & vbCrLf & "nothing" & vbCrLf & "something" & vbCrLf & "nothing"
           
           Application.Wait (Now + TimeValue("0:00:5"))

        Exit For
        End If
    Next t

    
    Set elems = .getElementsByTagName("input")

    For Each e In elems 'to click the go button

        If (e.getAttribute("name") = "btnGo") Then
        
            Application.Wait (Now + TimeValue("0:00:5"))
            
            e.Click
            
            Do While IE.Busy: DoEvents: Loop

            Exit For
        End If

    Next e
    
     For Each t In someThing ' to get the lat longs
        If (t.getAttribute("name") = "txtGetOK") Then
            Do While IE.Busy: DoEvents: Loop
            Application.Wait (Now + TimeValue("0:00:5"))
            str = t.Value
            Exit For
        End If
        Next t

End With
Do While IE.Busy: DoEvents: Loop
Range("A1") = t.Value


IE.Quit


'Debug.Print (InStr(1, str, vbCrLf, vbBinaryCompare))

'Debug.Print (InStr(1, str, vbTab, vbBinaryCompare))
StringArray = Split(str, vbTab)
Cells(1, 1) = StringArray(6)
Cells(1, 2) = StringArray(7)
Cells(1, 3) = StringArray(8)

End Sub

'IE.Document.getElementById("Form_Attempts-1372643500")(0).Value = "1"
' IE.Document.getElementById("submit-1-1372643500")(0).Click

