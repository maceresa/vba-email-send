'Add reference to mscorlib.tlb
Global listaMails As String
Global tipoContratacion As String
Global numeroContratacion As String
Global textoMail As String
Global grupoMails As String
Global folderName As String
Global mailList As New ArrayList
Global strPath As String

Global mailsColumn As Integer

Sub GetConfigValues()

    Worksheets("MACROS").Activate
    strPath = Cells(3, 2).Value
    numeroContratacion = Cells(7, 2).Value
    tipoContratacion = Cells(4, 2).Value
    textoMail = Cells(6, 2).Value
    mailsColumn = Cells(5, 2).Value
    grupoMails = Cells(8, 2).Value
    folderName = Cells(9, 2).Value
   
End Sub

Sub BuscarMails()

    Dim Cell As Range
    Dim rango As String
    
    Sheets(grupoMails).Activate
    For Each Cell In ActiveSheet.Range(Cells(2, mailsColumn), Cells(600, mailsColumn))
        If Not Cell.Value = "" Then
            mailList.Add Cell.Value
        End If
    Next Cell

End Sub

Sub SendEmail()

    Dim objOutlook As Object
    Dim objMail As Object
    Dim strTo As String
    Dim strSubject As String
    Dim strBody As String
    Dim item As Integer
    Dim listas As Double
    Dim max As Integer
    Dim attachmentPath As String

    max = 90
    
    GetConfigValues
    BuscarMails
    
    If Len(numeroContratacion) = 1 Then
        numeroContratacion = "000" & numeroContratacion
    End If
    If Len(numeroContratacion) = 2 Then
        numeroContratacion = "00" & numeroContratacion
    End If
    If Len(numeroContratacion) = 3 Then
        numeroContratacion = "0" & numeroContratacion
    End If
    
    'Set email properties
    strSubject = tipoContratacion & " " & numeroContratacion & " - Solicitud de Presupuesto"
    
    attachmentPath = strPath & folderName & "\" & tipoContratacion & " " & numeroContratacion & " - Pliego.pdf"
    
    Dim total As Integer
    total = mailList.Count
    listas = WorksheetFunction.RoundUp(mailList.Count / 90, 0)
    
    If mailList.Count > 0 Then
        item = 0
        'armo lista
        For J = 1 To listas
            If mailList.Count < max Then
                max = mailList.Count
            End If
            For N = item To max - 1
                strTo = strTo & mailList.item(N) & "; "
            Next N
        
            'mando mail
            Set objOutlook = CreateObject("Outlook.Application")
            Set objMail = objOutlook.CreateItem(0)
            
            With objMail
                .BCC = strTo
                .Subject = strSubject
                .Body = textoMail
                .Attachments.Add attachmentPath
                .Send
            End With
        
            'Clean up
            Set objMail = Nothing
            Set objOutlook = Nothing
            strTo = ""
                    
            max = 90 * (J + 1)
            item = 90 * J
        Next J
        
        mailList.Clear
        Sheets("MACROS").Activate
        MsgBox ("Envío Exitoso")
        
    Else
        MsgBox ("No se encontraron mails para enviar")
    End If

End Sub

