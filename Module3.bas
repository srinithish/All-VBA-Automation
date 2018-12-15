Attribute VB_Name = "Module3"
Sub ForAttendance()
Dim str As Variant
Dim IE As Object
Set IE = New InternetExplorer
IE.Visible = True
IE.Navigate "https://hris.excelityglobal.com/embrace/jsp/login.jsp"

Do While IE.Busy: DoEvents: Loop

With IE.Document

   
           
          
    Set elems = .getElementsByTagName("input")

    For Each e In elems 'to click the go button

        If (e.getAttribute("name") = "txtLoginId1") Then

            Application.Wait (Now + TimeValue("0:00:5"))

            e.Value = "5268113"

            Do While IE.Busy: DoEvents: Loop

'            Exit For
        End If
        
        If (e.getAttribute("name") = "txtPassword") Then
        
            e.Value = "kona02101993"
            Do While IE.Busy: DoEvents: Loop
            
'            Exit For
        End If
        
        If (e.getAttribute("name") = "txtCorporation1") Then
        
            e.Value = "aig"
            Do While IE.Busy: DoEvents: Loop
            
'          Exit For
        End If
        
        If (e.getAttribute("name") = "logOn") Then
        
            e.Click
            Do While IE.Busy: DoEvents: Loop
            
          Exit For
        End If
        

    Next e
    
    Application.Wait (Now + TimeValue("0:00:5"))
    
    IE.Navigate "https://hris.excelityglobal.com/embrace/servlet/controller?module=OnlineAttendance&screen=CreateRegularizeAttendance&action=View&empoid=010204135175147641614509600940&createFlag=Y"

    
'    Set someThing = .getElementsByTagName("HREF")
'    For Each x In someThing
'        Debug.Print ("hereh")
'        If someThing.href = "/embrace/servlet/controller?module=OnlineAttendance&screen=CreateRegularizeAttendance&action=View&empoid=010204135175147641614509600940&createFlag=Y" Then
'
'            'Application.Wait (Now + TimeValue("0:00:5"))
'
'            someThing.Click
'
'            Do While IE.Busy: DoEvents: Loop
'
'            Exit For
'        End If
'    Next x
    
End With

'     For Each t In someThing ' to get the lat longs
'        If (t.getAttribute("name") = "txtGetOK") Then
'            Do While IE.Busy: DoEvents: Loop
'            Application.Wait (Now + TimeValue("0:00:5"))
'            str = t.Value
'            Exit For
'        End If
'        Next t


Do While IE.Busy: DoEvents: Loop
'Range("A1") = t.Value


Set IE = Nothing

'Debug.Print (InStr(1, str, vbCrLf, vbBinaryCompare))

End Sub

'IE.Document.getElementById("Form_Attempts-1372643500")(0).Value = "1"
' IE.Document.getElementById("submit-1-1372643500")(0).Click


