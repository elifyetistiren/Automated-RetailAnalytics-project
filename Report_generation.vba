Option Explicit

Sub Generate_Report1()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call InitializeVariables
Call Connect_Database

Dim sql_query As String
Dim rp_week As Date, rp_week_end As Date
Dim rstDatabase As ADODB.Recordset
Dim row_data As Integer
Dim folder_save As String, file_save As String, full_save As String
Dim strTo As String, strCc As String, strSubject As String, strBody As String

rp_week = shtMain.Range("M10")
rp_week_end = rp_week + 7

With shtRp1
    .Range("I8:J12").ClearContents
    .Range("I15:J17").ClearContents
    .Range("I20:J24").ClearContents
    
    'Data Table 1
    sql_query = "SELECT tbCustomer.customer_name, SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity) as Total_Sales " & _
                "FROM tbSales INNER JOIN tbCustomer " & _
                "ON tbSales.client_id = tbCustomer.customer_id " & _
                "WHERE tbSales.sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ") " & _
                "GROUP BY tbCustomer.customer_name"
    
    Set rstDatabase = New Recordset
    row_data = 8
    
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        Do While Not .EOF
            shtRp1.Cells(row_data, 9) = .Fields("customer_name")
            shtRp1.Cells(row_data, 10) = .Fields("Total_Sales")
            row_data = row_data + 1
            .MoveNext
        Loop
    End With

    'Data Table 2
    sql_query = "SELECT TOP 3 tbShops.shop_location, SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity) as Total_Sales " & _
                "FROM tbSales INNER JOIN tbShops " & _
                "ON tbSales.shop_id = tbShops.shop_id " & _
                "WHERE tbSales.sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ") " & _
                "GROUP BY tbShops.shop_location " & _
                "ORDER BY SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity)"

    Set rstDatabase = New Recordset
    row_data = 15
    
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        Do While Not .EOF
            shtRp1.Cells(row_data, 9) = .Fields("shop_location")
            shtRp1.Cells(row_data, 10) = .Fields("Total_Sales")
            row_data = row_data + 1
            .MoveNext
        Loop
    End With

    'Data Table 3
    sql_query = "SELECT tbShops.shop_name, AVG(tbSales.sales_discount) as Avg_Discount " & _
                "FROM tbSales INNER JOIN tbShops " & _
                "ON tbSales.shop_id = tbShops.shop_id " & _
                "WHERE (tbSales.sales_status = 'Paid' AND tbSales.sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ")) " & _
                "GROUP BY tbShops.shop_name"

    Set rstDatabase = New Recordset
    row_data = 20
    
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        Do While Not .EOF
            shtRp1.Cells(row_data, 9) = .Fields("shop_name")
            shtRp1.Cells(row_data, 10) = .Fields("Avg_Discount")
            row_data = row_data + 1
            .MoveNext
        Loop
    End With
    
.Range("J32") = rp_week
Application.Calculate

folder_save = "C:\Users\amand\OneDrive\Documents\Elif dashboard project\Github uploaded project\Report1\"
file_save = "Sales Report - " & Format(rp_week, "yyyymmdd") & ".xlsx"
full_save = folder_save & file_save

.Copy
ActiveWorkbook.Sheets(1).Name = "Data"
ActiveWorkbook.SaveAs Filename:=full_save
ActiveWorkbook.Close

strTo = "sales@winterforecasting.com"
strCc = "finance@winterforecasting.com"
strSubject = "Weekly Sales Report - " & Format(rp_week, "yyyymmdd")
strBody = "<BODY style=font-family:Calibri>Dear Sales team, <p>Please find attached your weekly report.<p>Kind regards,<p><b>Finance Department</b></BODY>"
Call SendEmail(strTo, strCc, strSubject, strBody, full_save)

End With

Call Disconnect_Database

End Sub


Sub Generate_Report_2()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call InitializeVariables
Call Connect_Database

Dim sql_query As String
Dim rp_week As Date, rp_week_end As Date
Dim rstDatabase As ADODB.Recordset
Dim row_data As Integer
Dim folder_save As String, file_save As String, full_save As String
Dim strTo As String, strCc As String, strSubject As String, strBody As String

rp_week = shtMain.Range("M10")
rp_week_end = rp_week + 7

With shtRp2

    .Range("P7:Q11").ClearContents
    .Range("P17:Q25").ClearContents
    .Range("P31:Q35").ClearContents
    .Range("P41:R45").ClearContents
    
    'Table 1: Returned Product Sales Value
    sql_query = "SELECT tbShops.shop_name, SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity) as Total_Sales " & _
                "FROM tbSales INNER JOIN tbShops ON tbSales.shop_id = tbShops.shop_id " & _
                "WHERE (tbSales.sales_status = 'Returned' AND tbSales.sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ")) " & _
                "GROUP BY tbShops.shop_name " & _
                "ORDER BY SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity) ASC"
 
    row_data = 7
    
    Set rstDatabase = New Recordset
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        
        Do While Not .EOF
            shtRp2.Cells(row_data, 16) = .Fields("shop_name")
            shtRp2.Cells(row_data, 17) = -1 * .Fields("Total_Sales")
            row_data = row_data + 1
            .MoveNext
        Loop
        
    End With
    
    'Table 2: Sales Count per Opening Hour
    sql_query = "SELECT DatePart('h', sales_date) AS HourlySales, COUNT(sales_id) as Transaction_Count " & _
                "FROM tbSales " & _
                "WHERE ((sales_status IN ('Paid', 'Reserved')) AND (sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & "))) " & _
                "GROUP BY DatePart('h', sales_date)"
 
    row_data = 17
    
    Set rstDatabase = New Recordset
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        
        Do While Not .EOF
            shtRp2.Cells(row_data, 16) = .Fields("HourlySales")
            shtRp2.Cells(row_data, 17) = .Fields("Transaction_Count")
            row_data = row_data + 1
            .MoveNext
        Loop
        
    End With
    
    'Table 3: Sales Split per Location
    sql_query = "SELECT tbShops.shop_location, SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity) " & _
                "/ (SELECT SUM(sales_price*(1-sales_discount)*sales_quantity) FROM tbSales WHERE sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ")) as Total_Sales " & _
                "FROM tbSales INNER JOIN tbShops ON tbSales.shop_id = tbShops.shop_id " & _
                "WHERE tbSales.sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ") " & _
                "GROUP BY tbShops.shop_location " & _
                "ORDER BY SUM(tbSales.sales_price*(1-tbSales.sales_discount)*tbSales.sales_quantity)"

    row_data = 31
    
    Set rstDatabase = New Recordset
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        
        Do While Not .EOF
            shtRp2.Cells(row_data, 16) = .Fields("shop_location")
            shtRp2.Cells(row_data, 17) = .Fields("Total_Sales")
            row_data = row_data + 1
            .MoveNext
        Loop
        
    End With
    
    'Table 4: Weekly Sales per Shop
    sql_query = "SELECT tbShops.shop_name, tbTempResult.Total_Sales, tbTempResult.target_value " & _
                "FROM tbShops INNER JOIN " & _
                "(SELECT tbTempSales.shop_id, tbTempSales.Total_Sales, tbTempPerf.target_value " & _
                "FROM " & _
                    "(SELECT shop_id, SUM(sales_price*(1-sales_discount)*sales_quantity) as Total_Sales " & _
                    "FROM tbSales " & _
                    "WHERE sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ") " & _
                    "GROUP BY shop_id) as tbTempSales " & _
                "INNER Join " & _
                    "(SELECT shop_id, target_value " & _
                    "FROM tbPerformance " & _
                    "WHERE target_week = CDATE(" & CDbl(rp_week) & ")) as tbTempPerf " & _
                    "ON tbTempSales.shop_id = tbTempPerf.shop_id) as tbTempResult " & _
                "ON tbShops.shop_id = tbTempResult.shop_id " & _
                "ORDER BY tbTempResult.target_value ASC"

    row_data = 41
    
    Set rstDatabase = New Recordset
    With rstDatabase
        .Open Source:=sql_query, ActiveConnection:=cnnDatabase
        
        Do While Not .EOF
            shtRp2.Cells(row_data, 16) = .Fields("shop_name")
            shtRp2.Cells(row_data, 17) = .Fields("Total_Sales")
            shtRp2.Cells(row_data, 18) = .Fields("target_value")
            row_data = row_data + 1
            .MoveNext
        Loop
        
    End With
    
    .Range("I6") = rp_week
    Application.Calculate
    
    folder_save = "C:\Users\amand\OneDrive\Documents\Elif dashboard project\Github uploaded project\Report2\"
    file_save = "CEO Report - " & Format(rp_week, "yyyymmdd") & ".pdf"
    full_save = folder_save & file_save
    
    .ExportAsFixedFormat xlTypePDF, Filename:=full_save, ignoreprintareas:=False
    
    strTo = .Range("F54")
    strCc = .Range("F55")
    strSubject = "Weekly CEO Report - " & Format(rp_week, "yyyymmdd")
    strBody = .Range("F56")
    
    Call SendEmail(strTo, strCc, strSubject, strBody, full_save)
    
End With

Call Disconnect_Database

End Sub

Sub Generate_Report_3()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call InitializeVariables
Call Connect_Database

Dim sql_query As String
Dim rp_week As Date, rp_week_end As Date
Dim rstDatabase As ADODB.Recordset
Dim folder_save As String, file_save As String, full_save As String
Dim strTo As String, strCc As String, strSubject As String, strBody As String
Dim strTarget As String

rp_week = shtMain.Range("M10")
rp_week_end = rp_week + 7

sql_query = "SELECT tbShops.shop_name, tbShops.shop_manager, tbShops.shop_contact, tbTempResult.Total_Sales, tbTempResult.target_value " & _
            "FROM tbShops INNER JOIN " & _
            "(SELECT tbTempSales.shop_id, tbTempSales.Total_Sales, tbTempPerf.target_value " & _
            "FROM " & _
                "(SELECT shop_id, SUM(sales_price*(1-sales_discount)*sales_quantity) as Total_Sales " & _
                "FROM tbSales " & _
                "WHERE sales_date BETWEEN CDATE(" & CDbl(rp_week) & ") AND CDATE(" & CDbl(rp_week_end) & ") " & _
                "GROUP BY shop_id) as tbTempSales " & _
            "INNER Join " & _
                "(SELECT shop_id, target_value " & _
                "FROM tbPerformance " & _
                "WHERE target_week = CDATE(" & CDbl(rp_week) & ")) as tbTempPerf " & _
                "ON tbTempSales.shop_id = tbTempPerf.shop_id) as tbTempResult " & _
            "ON tbShops.shop_id = tbTempResult.shop_id " & _
            "ORDER BY tbTempResult.target_value ASC"

Set rstDatabase = New Recordset
With rstDatabase
    .Open Source:=sql_query, ActiveConnection:=cnnDatabase
    
    strCc = "finance@winterforecasting.com"
    
    Do While Not .EOF
        strTo = .Fields("shop_contact")
        strSubject = .Fields("shop_name") & " Weekly Report - " & Format(rp_week, "yyyymmdd")
                
        If .Fields("Total_Sales") >= .Fields("target_value") Then
            strTarget = "Congrats, you have reached your target!"
        Else
            strTarget = "Seems you didn't reach your target, you will do better next week!"
        End If
        
        strBody = "<BODY style=font-family:Calibri>Dear " & .Fields("shop_manager") & "," & _
                    "<p>Please find your weekly results below." & _
                    "<ul>" & _
                        "<li>Target for the week: " & Format(.Fields("target_value"), "Standard") & "</li>" & _
                        "<li>Result for the week: " & Format(.Fields("Total_Sales"), "Standard") & "</li>" & _
                    "</ul>" & _
                    "<p>" & strTarget & _
                    "<p>Kind regards,<p><b>Finance Department</b>" & _
                    "</BODY>"
        
        Call SendEmail(strTo, strCc, strSubject, strBody)
        .MoveNext
    Loop
    
End With


Call Disconnect_Database



End Sub


Public Sub SendEmail(strTo As String, strCc As String, strSubject As String, strBody As String, Optional strAttach As String)

Dim olApp As Object
Dim olMail As Object

Set olApp = CreateObject("Outlook.Application")
Set olMail = olApp.CreateItem(olMailItem)

With olMail
    .To = strTo
    .Cc = strCc
    .Subject = strSubject
    .HTMLBody = strBody
    If strAttach <> "" Then
        .attachments.Add (strAttach)
    End If
    .display
End With

End Sub















