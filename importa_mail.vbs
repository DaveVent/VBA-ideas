'SUB PER IMPORTARE MAIL DA SOTTOCARTELLE OUTLOOK IN EXCEL (OGNI SOTTOCARTELLA DEVE AVERE UN FOGLIO SPECIFICO NEL FILE EXCEL CON IL PROPRIO NOME)

Sub importa_mail()

'Eliminare updates e messaggi di avviso dell'applicazione
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
'Dichiarare le variabili
Dim OutlookApp As Object, olNs As Object, olFolder As Object, wk As Worksheet

'Gestione dell'errore
On Error GoTo ERRORE

'Inizializzare le variabili di Outlook (MAPI = protocollo obbligatorio per l'accesso alla mail)
Set OutlookApp = New Outlook.Application
Set olNs = OutlookApp.GetNamespace("MAPI")

'Inizializzare ciclo che iteri su ogni foglio della cartella attiva
For Each wk In ActiveWorkbook.Worksheets

'Attivare il foglio
wk.Activate

'Stabilire la connessione alla cartella di Outlook che ha lo stesso nome del foglio attivo
    Set olFolder = olNs.GetDefaultFolder(olFolderInbox).Folders(ActiveSheet.Name)
    
'Se la cartella di Outlook è vuota, visualizzare un messaggio informativo
If olFolder.Items.Count = 0 Then
    MsgBox "nessuna mail da scaricare nella cartella " & ActiveSheet.Name
End If

    'Inizializzare ciclo che iteri su ogni elemento della cartella
    For x = 1 To olFolder.Items.Count
    
        'ISTRUZIONE CONDIZIONALE OPZIONALE PER ESTRARRE SOLO LE MAIL OLTRE UNA CERTA DATA
        'If Month(olFolder.Items.Item(x).ReceivedTime) = Month(Now) Then
        
        'Estrarre data di ricezione, mittente, oggetto e messaggio di ogni mail
        With ActiveSheet
            .Range("a" & 1 + x).Value = olFolder.Items.Item(x).ReceivedTime
            .Range("b" & 1 + x).Value = olFolder.Items.Item(x).SenderName
            .Range("c" & 1 + x).Value = olFolder.Items.Item(x).Subject
            .Range("d" & 1 + x).Value = olFolder.Items.Item(x).Body
        End With
        
        'End If
        
    Next x

'Eliminare l' a capo
Columns("A:D").Select
With Selection
    .WrapText = False
End With
            
    'OPZIONALE PER ELIMINARE LE RIGHE VUOTE
'Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
'Autosize delle colonne
Columns("A:C").AutoFit


Next wk

Exit Sub
    
'Messaggio in caso di errore
ERRORE:
    MsgBox "Directory delle cartelle non corretta", vbCritical
    
End Sub
