## Excel Mail Marge
if want to save individual file with excel mail marge the 
# Create a folder name C:\MergedPDFs
## Now code is 
```vbnet
Sub ExportEachCustomerToPDF()
    Dim dataDoc As Document
    Dim tempDoc As Document
    Dim savePath As String
    Dim pdfPath As String
    Dim i As Integer
    Dim totalRecords As Integer
    Dim customerName As String

    savePath = "C:\MergedPDFs\"
    If Dir(savePath, vbDirectory) = "" Then MkDir savePath

    Set dataDoc = ActiveDocument
    totalRecords = dataDoc.MailMerge.DataSource.RecordCount

    For i = 1 To totalRecords
        With dataDoc.MailMerge
            .Destination = wdSendToNewDocument
            .SuppressBlankLines = True
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            .Execute Pause:=False
        End With

        Set tempDoc = ActiveDocument

        ' Get customer name from 1st line
        On Error Resume Next
        customerName = Trim(tempDoc.Paragraphs(1).Range.Words(2))
        If customerName = "" Then customerName = "Customer_" & i
        On Error GoTo 0

        ' Clean name for filename
        customerName = Replace(customerName, "\", "")
        customerName = Replace(customerName, "/", "")
        customerName = Replace(customerName, ":", "")
        customerName = Replace(customerName, "*", "")
        customerName = Replace(customerName, "?", "")
        customerName = Replace(customerName, """", "")
        customerName = Replace(customerName, "<", "")
        customerName = Replace(customerName, ">", "")
        customerName = Replace(customerName, "|", "")
        customerName = Trim(customerName)

        pdfPath = savePath & customerName & ".pdf"
        tempDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
        tempDoc.Close False
    Next i

    MsgBox "? All PDFs saved to: " & savePath
End Sub

```
# Now if you want to send mail with word file then 
```vbnet
Sub ExportEachCustomerToPDFAndSendEmail()
    Dim dataDoc As Document
    Dim tempDoc As Document
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim savePath As String
    Dim pdfPath As String
    Dim customerName As String
    Dim customerEmail As String
    Dim i As Integer
    Dim totalRecords As Integer

    savePath = "C:\MergedPDFs\"
    If Dir(savePath, vbDirectory) = "" Then MkDir savePath

    Set dataDoc = ActiveDocument
    totalRecords = dataDoc.MailMerge.DataSource.RecordCount

    ' Start Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If OutlookApp Is Nothing Then
        MsgBox "Outlook not found.", vbCritical
        Exit Sub
    End If

    For i = 1 To totalRecords
        With dataDoc.MailMerge
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            
            ' Get email and name from DataFields BEFORE merge execution
            customerEmail = .DataSource.DataFields("Email").Value
            customerName = .DataSource.DataFields("Name").Value
            
            .Destination = wdSendToNewDocument
            .SuppressBlankLines = True
            .Execute Pause:=False
        End With

        Set tempDoc = ActiveDocument

        ' Clean filename to avoid illegal characters
        customerName = Replace(customerName, "\", "")
        customerName = Replace(customerName, "/", "")
        customerName = Replace(customerName, ":", "")
        customerName = Replace(customerName, "*", "")
        customerName = Replace(customerName, "?", "")
        customerName = Replace(customerName, """", "")
        customerName = Replace(customerName, "<", "")
        customerName = Replace(customerName, ">", "")
        customerName = Replace(customerName, "|", "")
        customerName = Trim(customerName)

        pdfPath = savePath & customerName & ".pdf"
        tempDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
        tempDoc.Close False

        ' Validate email before sending
        If InStr(customerEmail, "@") > 0 Then
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = customerEmail
                .Subject = "Your Order Confirmation"
                .Body = "Dear " & customerName & "," & vbCrLf & vbCrLf & _
                        "Thank you for your order. Please find the attached order summary PDF." & vbCrLf & vbCrLf & _
                        "Best regards," & vbCrLf & "Your Company"
                .Attachments.Add pdfPath
                .Send  ' Change to .Display if you want to preview before sending
            End With
        Else
            MsgBox "Invalid email for record #" & i & ": " & customerEmail, vbExclamation
        End If
    Next i

    MsgBox "✅ All emails with PDFs sent!", vbInformation
End Sub
```
