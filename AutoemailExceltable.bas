Attribute VB_Name = "autoemail"
Sub SendEmails()

    ' Set up Outlook object
    Dim OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    
    ' Get worksheet with recipient information
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("pivot")
    
    'Get signature
    Dim OApp As Object, OMail As Object, signature As String
    Set OApp = CreateObject("Outlook.Application")
    Set OMail = OApp.CreateItem(0)
    With OMail
        .Display
    End With
    signature = OMail.HTMLBody
    
    ' Get last row of data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Create dictionary to store unique email addresses and their corresponding body content
    Dim emailDict As Object
    Set emailDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through recipients and add their body content to dictionary
    Dim i As Long
    For i = 22 To lastRow
        Dim email As String
        email = ws.Range("J" & i).Value
        If email = "" Then
        ' Skip this row and go to the next one
        GoTo SkipRow
        End If
        If emailDict.Exists(email) Then
            emailDict(email) = emailDict(email) & "<tr>" & _
                "<td>" & ws.Range("A" & i).Value & "</td>" & _
                "<td>" & ws.Range("B" & i).Value & "</td>" & _
                "<td>" & ws.Range("C" & i).Value & "</td>" & _
                "<td>" & ws.Range("D" & i).Value & "</td>" & _
                "<td>" & ws.Range("E" & i).Value & "</td>" & _
                "<td>" & ws.Range("F" & i).Value & "</td>" & _
                "<td>" & ws.Range("G" & i).Value & "</td>" & _
                "<td>" & ws.Range("H" & i).Value & "</td>" & _
                "<td>" & ws.Range("I" & i).Value & "</td>" & _
            "</tr>"
        Else
            emailDict.Add email, "<table border='1' style='font-size:20px'>" & _
                "<tr>" & _
                    "<th>" & ws.Range("A21").Value & "</th>" & _
                    "<th>" & ws.Range("B21").Value & "</th>" & _
                    "<th>" & ws.Range("C21").Value & "</th>" & _
                    "<th>" & ws.Range("D21").Value & "</th>" & _
                    "<th>" & ws.Range("E21").Value & "</th>" & _
                    "<th>" & ws.Range("F21").Value & "</th>" & _
                    "<th>" & ws.Range("G21").Value & "</th>" & _
                    "<th>" & ws.Range("H21").Value & "</th>" & _
                     "<th>" & ws.Range("I21").Value & "</th>" & _
                "</tr>"
             emailDict(email) = emailDict(email) & "<tr>" & _
                "<td>" & ws.Range("A" & i).Value & "</td>" & _
                "<td>" & ws.Range("B" & i).Value & "</td>" & _
                "<td>" & ws.Range("C" & i).Value & "</td>" & _
                "<td>" & ws.Range("D" & i).Value & "</td>" & _
                "<td>" & ws.Range("E" & i).Value & "</td>" & _
                "<td>" & ws.Range("F" & i).Value & "</td>" & _
                "<td>" & ws.Range("G" & i).Value & "</td>" & _
                "<td>" & ws.Range("H" & i).Value & "</td>" & _
                 "<td>" & ws.Range("I" & i).Value & "</td>" & _
            "</tr>"
        End If
SkipRow:
    Next i
    

    
   ' Loop through unique email addresses in dictionary and send email
    Dim emailKey As Variant
    For Each emailKey In emailDict.Keys
    
    ' Set email body and subject
    Dim body As String
    body = ws.Range("H4").Value & "<br><br><table>" & emailDict(emailKey) & "</table>" & signature



    Dim subject As String
    subject = ws.Range("I1").Value
    
    ' Send email to recipient
    Dim OutMail As Object
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = emailKey
        .subject = ws.Range("H2").Value
        .HTMLBody = body
        .Send
    End With
    Set OutMail = Nothing
    
    Next emailKey

    
    ' Clean up
    Set OutApp = Nothing
    Set ws = Nothing
    
    MsgBox "All emails sent."
    
End Sub














