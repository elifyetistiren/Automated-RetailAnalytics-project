Option Explicit

Public cnnDatabase As ADODB.Connection
Public blnIsConnected As Boolean

Public Sub Connect_Database()

blnIsConnected = False

On Error GoTo errHandling

Set cnnDatabase = New ADODB.Connection

With cnnDatabase
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    'IMPORTANT: Replace the "..." below by the main location where the Data Tree has been saved
    .ConnectionString = "Data Source = C:\Users\amand\OneDrive\Documents\Elif dashboard project\Github uploaded project\Finance_DB.accdb"
    .Properties("Jet OLEDB:Database Password") = ""
    .Open
End With

blnIsConnected = True

Exit Sub

errHandling:
MsgBox "Connection to database failed!", vbCritical, "Error!"

End Sub


Public Sub Disconnect_Database()

cnnDatabase.Close
Set cnnDatabase = Nothing
blnIsConnected = False

End Sub

