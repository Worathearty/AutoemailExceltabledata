Attribute VB_Name = "autoemail"
Sub SAutoEmails()
    
    ' Set up Outlook object
    Dim OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    
    ' Prompt for email subject, body, and attachments
    Dim subject As String
    Dim body As String
    subject = ThisWorkbook.Sheets("Recipients").Range("I2").Value
    body = "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I5").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I6").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I7").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I8").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I9").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I10").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I11").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I12").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I13").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I14").Value & "</p>" & _
       "<p style='font-size:20px;'>" & ThisWorkbook.Sheets("Recipients").Range("I15").Value & "</p>"

    
    Dim ccAddress As String
    ccAddress = InputBox("Enter the CC address:")
        
        ' Prompt user to select attachment files
    Dim attachmentPaths As Variant
    attachmentPaths = Application.GetOpenFilename(FileFilter:="All Files (*.*),*.*", MultiSelect:=True)
    
    Dim attachmentNames As String
    If IsArray(attachmentPaths) Then
        For Each attachmentPath In attachmentPaths
            attachmentNames = attachmentNames & vbCrLf & Mid(attachmentPath, InStrRev(attachmentPath, "\") + 1)
        Next attachmentPath
    Else
        attachmentNames = "No attachments"
    End If
    
    
    Dim confirmBody As String
    confirmBody = Replace(body, "<p>", "")
    confirmBody = Replace(confirmBody, "</p>", "")
    
    
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox("Do you want to send the following email?" & vbCrLf & _
                       "Subject: " & subject & vbCrLf & _
                       "Body: " & vbCrLf & confirmBody & vbCrLf & _
                       "Attachments: " & attachmentNames, vbOKCancel)


    If confirmResult = vbCancel Then
        Exit Sub
    End If
    
    ' Loop through recipients in Excel worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Recipients")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    
    
    
    Dim OApp As Object, OMail As Object, signature As String
    Set OApp = CreateObject("Outlook.Application")
    Set OMail = OApp.CreateItem(0)
    With OMail
        .Display
    End With
    signature = OMail.HTMLBody

    
    Dim recipient As String
    Dim i As Long
    For i = 2 To lastRow
        recipient = ws.Range("A" & i).Value
        
        ' Send email to recipient
        Dim OutMail As Object
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .To = recipient
            .CC = ccAddress
            .subject = subject
            .HTMLBody = "<p style='font-size:20px;'>เรียน " & ws.Range("B" & i).Value & "</p>" & _
            "<p>" & body & "</p>" & _
            signature
            
                    ' Attach files to email
        If IsArray(attachmentPaths) Then
            For Each attachmentPath In attachmentPaths
                .attachments.Add attachmentPath
            Next attachmentPath
        End If

                       
            ' Send email
            .Send
            
            ' Add status to column C
            ws.Range("C" & i).Value = "Sent"
        End With
        Set OutMail = Nothing
    Next i
    
    ' Clean up
    Set OutApp = Nothing
    Set ws = Nothing
    
    MsgBox "All emails sent."
    
End Sub













