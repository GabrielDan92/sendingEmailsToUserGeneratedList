Public raw As Worksheet, dest As Worksheet, criteria As Worksheet, final As Worksheet
Public rng As Range, rngData As Range, rngCriteria As Range, rngOutput As Range, month As String

Sub email(ByVal month As String)

'the main module is called using a userform with a dropdown list where the end user can select the month he/she wants

Dim lastRow As Long, lastRowFinal As Long, i As Long
Dim arr As Variant, wb As Workbook, pass As String
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Set criteria = Sheet1

'find out how many emails have been added by the user in the specified column
'in this case is column "G", but it can be whatever column you and the end user decide
lastRow = criteria.Range("G1", criteria.Range("G2").End(xlDown)).Count

'loop through each row and call the sendEmail() function that generates an email and sends it to the email address found
For i = 2 To lastRow
    Call sendEmail(i, month)
Next i


Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Function sendEmail(ByVal j As Integer, ByVal month As String)

    Dim outlookApp As New Outlook.Application
    Dim newEmail As Outlook.MailItem
    Set newEmail = outlookApp.CreateItem(olMailItem)

    With criteria
        emailAddress = Application.WorksheetFunction.VLookup(.Range("G" & j).Value, .Range("J:K"), 2, 0)
        ContactName = .Range("G" & j).Value
        deadline = .Range("M2").Value
    End With
    
'    Debug.Print "email: " & emailAddress
'    Debug.Print "name: " & ContactName
'    Debug.Print "deadline: " & deadline

'create the email
    newEmail.To = emailAddress
    newEmail.CC = ""
    newEmail.SentOnBehalfOfName = ""
    newEmail.Subject = "Solicitare comentarii pentru luna " & month

'email's body
    newEmail.HTMLBody = "Buna ziua, <br><p>" _
    & "Pentru efectuarea procesului din luna aceasta, " _
    & "va rog sa adaugati in fisierul shared " _
    & "comentariile care lipsesc pentru facturile de pe numele dumneavoastra. <br><p>" _
    & " <br><p>" _
    & "Multumim, <br>" _
    & "Echipa Reporting"


'send the email
    'newEmail.display
    newEmail.Send

    Set outlookApp = Nothing

End Function
