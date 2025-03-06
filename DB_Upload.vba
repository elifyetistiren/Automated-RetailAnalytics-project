Option Explicit

Sub Upload_Performance()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim file_extension As String
Dim input_file As Variant
Dim wbData As Workbook
Dim shtData As Worksheet
Dim row_upload As Integer
Dim sql_query As String

Call InitializeVariables
Call Connect_Database

file_extension = "Excel Files (*.xlsx), *.xlsx"
input_file = Application.GetOpenFilename(filefilter:=file_extension, Title:="Please select a performance input file")

If input_file <> False Then
    'UPLOAD code
    Set wbData = Workbooks.Open(Filename:=input_file)
    Set shtData = wbData.Sheets("Data")
    
    row_upload = 2
    
    Do While Not IsEmpty(shtData.Cells(row_upload, 1))
        'Code to upload data in the database ...
        sql_query = "DELETE FROM tbPerformance WHERE target_id = " & shtData.Cells(row_upload, 1)
        cnnDatabase.Execute sql_query
        
        sql_query = "INSERT INTO tbPerformance (target_id, target_week, shop_id, target_value) VALUES (" & shtData.Cells(row_upload, 1) & ", " & CDbl(shtData.Cells(row_upload, 2)) & ", " & shtData.Cells(row_upload, 4) & ", " & shtData.Cells(row_upload, 8) & ")"
        cnnDatabase.Execute sql_query
        
        row_upload = row_upload + 1
    Loop
    
    wbData.Close
    MsgBox "File loaded in database!", vbInformation, "Microsoft Access Database"
    
Else
    MsgBox "No File selected, process will now end!", vbExclamation, "File selection error"
End If


Call Disconnect_Database

End Sub



Sub Upload_Sales()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim folder_path As String, file_extension As String, input_file As String, destination_path As String
Dim wbData As Workbook
Dim shtData As Worksheet
Dim row_upload As Integer, loaded_file_count As Integer
Dim rstDatabase As ADODB.Recordset

Call InitializeVariables
Call Connect_Database

folder_path = "C:\Users\amand\OneDrive\Documents\Elif dashboard project\Github uploaded project\Sales\"
destination_path = "C:\Users\amand\OneDrive\Documents\Elif dashboard project\Github uploaded project\Sales\Loaded\"

file_extension = "*.xlsx"
input_file = Dir(folder_path & file_extension)
loaded_file_count = 0

Set rstDatabase = New ADODB.Recordset

With rstDatabase

    .Open Source:="tbSales", ActiveConnection:=cnnDatabase, Locktype:=adLockOptimistic

    Do While input_file <> ""
        Set wbData = Workbooks.Open(Filename:=folder_path & input_file)
        Set shtData = wbData.Sheets("Data")
        
        row_upload = 2
        
        Do While Not IsEmpty(shtData.Cells(row_upload, 1))
            'Code to load the data into Access
            .AddNew
            .Fields("sales_id") = shtData.Cells(row_upload, 1)
            .Fields("sales_date") = shtData.Cells(row_upload, 2)
            .Fields("shop_id") = shtData.Cells(row_upload, 3)
            .Fields("product_id") = shtData.Cells(row_upload, 4)
            .Fields("client_id") = shtData.Cells(row_upload, 5)
            .Fields("sales_status") = shtData.Cells(row_upload, 6)
            .Fields("sales_quantity") = shtData.Cells(row_upload, 7)
            .Fields("sales_price") = shtData.Cells(row_upload, 8)
            .Fields("sales_discount") = shtData.Cells(row_upload, 9)
            .Update
            
            row_upload = row_upload + 1
        Loop
        
        wbData.Close
        Call Move_File(input_folder:=folder_path, destination_folder:=destination_path, input_file:=input_file)
        loaded_file_count = loaded_file_count + 1
        input_file = Dir
        
    Loop

End With

Call Disconnect_Database

If loaded_file_count > 0 Then
    MsgBox loaded_file_count & " input files loaded in database!", vbInformation, "Microsoft Access Database"
Else
    MsgBox "No input file loaded in database", vbExclamation, "Microsoft Access Database"
End If

End Sub

Private Sub Move_File(input_folder As String, destination_folder As String, input_file As String)

Dim FSO As Object

Set FSO = CreateObject("Scripting.filesystemobject")
FSO.Movefile Source:=input_folder & input_file, Destination:=destination_folder & input_file

End Sub
