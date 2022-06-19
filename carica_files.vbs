'SUB PER CONSOLIDARE DIVERSI FOGLI EXCEL SALVATI IN UNA DIRECTORY IN UN UNICO FILE

Sub carica_files()
  
  'dichiarare le variabili
Dim oFSO As Object, oFolder As Object, oFile As Object, i As Integer, urow As Long, wb_name As String, wb_file_name As String, directory As String
  
  'in caso di errore, andare alla sezione ERRORE
On Error GoTo ERRORE
  
  'inizializzare la directory
directory = InputBox("Copia il percorso dove hai salvato i file") & "\"
  'inizializzare l'oggetto FileSystem
Set oFSO = CreateObject("Scripting.FileSystemObject")
  'inizializzare l'accesso alla directory
Set oFolder = oFSO.GetFolder(directory)
  'salvare il nome della cartella aperta nella variabile wb_name
wb_name = ActiveWorkbook.Name

  'iterare sui file salvati nella directory
For Each oFile In oFolder.Files
    
      'disabilitare gli avvisi
    Application.DisplayAlerts = False
      'estrarre l'ultima riga della cartella aperta
    urow = Workbooks(wb_name).Sheets("Estrazione Dati").Cells(Rows.Count, "A").End(xlUp).Row
      'aprire il file salvato nella directory
    Workbooks.Open directory & oFile.Name
      'salvare il nome del file nella variabile wb_file_name
    wb_file_name = ActiveWorkbook.Name
      'copiare tutto il contenuto del file
    ActiveWorkbook.Sheets(1).Range("a1").CurrentRegion.Copy
      'attivare la cartella wb_name 
    Workbooks(wb_name).Activate
      'selezionare il primo foglio della cartella e incollare il contenuto del file
    Sheets(1).Range("a" & urow + 2).Select
    ActiveSheet.Paste
      'chiudere la cartella wb_file_name
    Workbooks(wb_file_name).Close savechanges:=False
    
      'riabilitare gli avvisi
    Application.DisplayAlerts = True
    
Next oFile

Exit Sub

'Gestione dell'errore
ERRORE:
  MsgBox "Directory non trovata!", vbCritical

End Sub
