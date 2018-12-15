Attribute VB_Name = "SendEmailFunction"
Sub SendEmail()
Dim OutlookApp As Object
Dim OutLookMailItem As Object
Set OutlookApp = CreateObject("Outlook.application")
Set OutLookMailItem = OutlookApp.CreateItem(0)

With OutLookMailItem
.Subject = "FYI"
.To = "nithish.kandagadla@gmail.com"
.Send
End With
End Sub
