Sub MailSender_Excel()
'
' Mailing Macro


    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim total As Integer
    Dim vacias As Integer
    Dim materiales As Integer
    Sheets("Menu").Select
    to1 = Range("L3").Value
    cc1 = Range("L7").Value
    sb = Range("L11").Value

 'Use the the Date Round up in 10min from the actual time
    vtdia = Format(CDate(Round(CDate(Time) * 1440 / 10, 0) * 10 / 1440), "hh:mm")

 'Enter the Outlook and copy the Signature from Outlook

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    

    On Error Resume Next
    
    Set OMail = OApp.CreateItem(0)
    With OutMail
        .Display
    End With
        Signature = OutMail.htmlbody 'Automatically insert the signature from Outlook
    
    With OutMail
        .To = to1
        .CC = cc1
        .BCC = ""
        .BodyFormat = olFormatHTML
        .Subject = sb
        .Font.Name = "Calibri"
        .htmlbody = "Body"
        .Display
        '.Send   'or use .Display     'If you want to send it automatically.
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub