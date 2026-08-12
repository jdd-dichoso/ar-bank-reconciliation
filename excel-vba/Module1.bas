Attribute VB_Name = "Module1"
Sub TestDataAccess()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, "A").End(xlUp).Row
    
    MsgBox "AR Data Records: " & lastRowAR - 1 & vbCrLf & _
           "Bank Data Records: " & lastRowBank - 1, _
           vbInformation, "Data Access Test"

End Sub


Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long

    Dim headerCell As Range
    
    Set headerCell = ws.Rows(1).Find( _
        What:=headerName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)
    
    If headerCell Is Nothing Then
        FindHeaderColumn = 0
    Else
        FindHeaderColumn = headerCell.Column
    End If

End Function

Sub TestHeaderMapping()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    
    Dim arInvoiceID As Long
    Dim arCustomer As Long
    Dim arAmount As Long
    Dim arCurrency As Long
    Dim arDate As Long
    
    Dim bankInvoiceID As Long
    Dim bankCustomer As Long
    Dim bankAmount As Long
    Dim bankCurrency As Long
    Dim bankDate As Long
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    
    'Find required AR columns
    arInvoiceID = FindHeaderColumn(wsAR, "Invoice_ID")
    arCustomer = FindHeaderColumn(wsAR, "Customer_Name")
    arAmount = FindHeaderColumn(wsAR, "Invoice_Amount")
    arCurrency = FindHeaderColumn(wsAR, "Currency")
    arDate = FindHeaderColumn(wsAR, "Cleared_Date_System")
    
    'Find required Bank columns
    bankInvoiceID = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    bankCustomer = FindHeaderColumn(wsBank, "Payer_Name")
    bankAmount = FindHeaderColumn(wsBank, "Amount_Received")
    bankCurrency = FindHeaderColumn(wsBank, "Currency")
    bankDate = FindHeaderColumn(wsBank, "Deposit_Date")
    
    MsgBox _
        "AR Data:" & vbCrLf & _
        "Invoice_ID: Column " & arInvoiceID & vbCrLf & _
        "Customer_Name: Column " & arCustomer & vbCrLf & _
        "Invoice_Amount: Column " & arAmount & vbCrLf & _
        "Currency: Column " & arCurrency & vbCrLf & _
        "Cleared_Date_System: Column " & arDate & vbCrLf & vbCrLf & _
        "Bank Data:" & vbCrLf & _
        "Reference_Invoice_ID: Column " & bankInvoiceID & vbCrLf & _
        "Payer_Name: Column " & bankCustomer & vbCrLf & _
        "Amount_Received: Column " & bankAmount & vbCrLf & _
        "Currency: Column " & bankCurrency & vbCrLf & _
        "Deposit_Date: Column " & bankDate, _
        vbInformation, "Header Mapping Test"

End Sub


Sub TestDataRead()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    
    Dim arInvoiceCol As Long
    Dim arCustomerCol As Long
    Dim arAmountCol As Long
    
    Dim bankInvoiceCol As Long
    Dim bankCustomerCol As Long
    Dim bankAmountCol As Long
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    
    Dim arData As Variant
    Dim bankData As Variant
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    
    'Find required columns
    arInvoiceCol = FindHeaderColumn(wsAR, "Invoice_ID")
    arCustomerCol = FindHeaderColumn(wsAR, "Customer_Name")
    arAmountCol = FindHeaderColumn(wsAR, "Invoice_Amount")
    
    bankInvoiceCol = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    bankCustomerCol = FindHeaderColumn(wsBank, "Payer_Name")
    bankAmountCol = FindHeaderColumn(wsBank, "Amount_Received")
    
    'Find last rows
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arInvoiceCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankInvoiceCol).End(xlUp).Row
    
    'Load data into memory
    arData = wsAR.Range(wsAR.Cells(2, arInvoiceCol), _
                        wsAR.Cells(lastRowAR, arAmountCol)).Value
    
    bankData = wsBank.Range(wsBank.Cells(2, bankInvoiceCol), _
                            wsBank.Cells(lastRowBank, bankAmountCol)).Value
    
    MsgBox _
        "Data successfully loaded into memory." & vbCrLf & vbCrLf & _
        "AR records loaded: " & UBound(arData, 1) & vbCrLf & _
        "Bank records loaded: " & UBound(bankData, 1), _
        vbInformation, "Data Read Test"

End Sub

Sub TestStandardization()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    
    Dim arInvoiceCol As Long
    Dim arCustomerCol As Long
    Dim arAmountCol As Long
    Dim arCurrencyCol As Long
    Dim arDateCol As Long
    
    Dim bankInvoiceCol As Long
    Dim bankCustomerCol As Long
    Dim bankAmountCol As Long
    Dim bankCurrencyCol As Long
    Dim bankDateCol As Long
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    
    Dim arStandardized() As Variant
    Dim bankStandardized() As Variant
    
    Dim i As Long
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    
    'Find AR columns
    arInvoiceCol = FindHeaderColumn(wsAR, "Invoice_ID")
    arCustomerCol = FindHeaderColumn(wsAR, "Customer_Name")
    arAmountCol = FindHeaderColumn(wsAR, "Invoice_Amount")
    arCurrencyCol = FindHeaderColumn(wsAR, "Currency")
    arDateCol = FindHeaderColumn(wsAR, "Cleared_Date_System")
    
    'Find Bank columns
    bankInvoiceCol = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    bankCustomerCol = FindHeaderColumn(wsBank, "Payer_Name")
    bankAmountCol = FindHeaderColumn(wsBank, "Amount_Received")
    bankCurrencyCol = FindHeaderColumn(wsBank, "Currency")
    bankDateCol = FindHeaderColumn(wsBank, "Deposit_Date")
    
    'Find last rows
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arInvoiceCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankInvoiceCol).End(xlUp).Row
    
    'Create standardized AR array
    ReDim arStandardized(1 To lastRowAR - 1, 1 To 6)
    
    For i = 2 To lastRowAR
        
        arStandardized(i - 1, 1) = wsAR.Cells(i, arInvoiceCol).Value
        arStandardized(i - 1, 2) = wsAR.Cells(i, arCustomerCol).Value
        arStandardized(i - 1, 3) = wsAR.Cells(i, arAmountCol).Value
        arStandardized(i - 1, 4) = wsAR.Cells(i, arCurrencyCol).Value
        arStandardized(i - 1, 5) = wsAR.Cells(i, arDateCol).Value
        arStandardized(i - 1, 6) = "AR Ledger"
        
    Next i
    
    'Create standardized Bank array
    ReDim bankStandardized(1 To lastRowBank - 1, 1 To 6)
    
    For i = 2 To lastRowBank
        
        bankStandardized(i - 1, 1) = wsBank.Cells(i, bankInvoiceCol).Value
        bankStandardized(i - 1, 2) = wsBank.Cells(i, bankCustomerCol).Value
        bankStandardized(i - 1, 3) = wsBank.Cells(i, bankAmountCol).Value
        bankStandardized(i - 1, 4) = wsBank.Cells(i, bankCurrencyCol).Value
        bankStandardized(i - 1, 5) = wsBank.Cells(i, bankDateCol).Value
        bankStandardized(i - 1, 6) = "Bank Remittance"
        
    Next i
    
    MsgBox _
        "Standardization successful." & vbCrLf & vbCrLf & _
        "AR records standardized: " & UBound(arStandardized, 1) & vbCrLf & _
        "Bank records standardized: " & UBound(bankStandardized, 1) & vbCrLf & vbCrLf & _
        "Fields:" & vbCrLf & _
        "1. recon_key" & vbCrLf & _
        "2. customer" & vbCrLf & _
        "3. amount" & vbCrLf & _
        "4. currency" & vbCrLf & _
        "5. transaction_date" & vbCrLf & _
        "6. source", _
        vbInformation, "Standardization Test"

End Sub

Sub TestGrouping()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    
    Dim arKeyCol As Long
    Dim bankKeyCol As Long
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    
    Dim arGroups As Object
    Dim bankGroups As Object
    
    Dim i As Long
    Dim key As String
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    
    arKeyCol = FindHeaderColumn(wsAR, "Invoice_ID")
    bankKeyCol = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arKeyCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankKeyCol).End(xlUp).Row
    
    Set arGroups = CreateObject("Scripting.Dictionary")
    Set bankGroups = CreateObject("Scripting.Dictionary")
    
    'Count AR records by recon_key
    For i = 2 To lastRowAR
        
        key = Trim(CStr(wsAR.Cells(i, arKeyCol).Value))
        
        If key <> "" Then
            
            If arGroups.Exists(key) Then
                arGroups(key) = arGroups(key) + 1
            Else
                arGroups.Add key, 1
            End If
            
        End If
        
    Next i
    
    'Count Bank records by recon_key
    For i = 2 To lastRowBank
        
        key = Trim(CStr(wsBank.Cells(i, bankKeyCol).Value))
        
        If key <> "" Then
            
            If bankGroups.Exists(key) Then
                bankGroups(key) = bankGroups(key) + 1
            Else
                bankGroups.Add key, 1
            End If
            
        End If
        
    Next i
    
    MsgBox _
        "Grouping successful." & vbCrLf & vbCrLf & _
        "Unique AR keys: " & arGroups.Count & vbCrLf & _
        "Unique Bank keys: " & bankGroups.Count, _
        vbInformation, "Grouping Test"

End Sub

Sub TestReconciliationUniverse()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    
    Dim arKeyCol As Long
    Dim bankKeyCol As Long
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    
    Dim allKeys As Object
    
    Dim i As Long
    Dim key As String
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    
    arKeyCol = FindHeaderColumn(wsAR, "Invoice_ID")
    bankKeyCol = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arKeyCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankKeyCol).End(xlUp).Row
    
    Set allKeys = CreateObject("Scripting.Dictionary")
    
    'Add AR keys
    For i = 2 To lastRowAR
        
        key = Trim(CStr(wsAR.Cells(i, arKeyCol).Value))
        
        If key <> "" Then
            If Not allKeys.Exists(key) Then
                allKeys.Add key, True
            End If
        End If
        
    Next i
    
    'Add Bank keys
    For i = 2 To lastRowBank
        
        key = Trim(CStr(wsBank.Cells(i, bankKeyCol).Value))
        
        If key <> "" Then
            If Not allKeys.Exists(key) Then
                allKeys.Add key, True
            End If
        End If
        
    Next i
    
    MsgBox _
        "Reconciliation universe created." & vbCrLf & vbCrLf & _
        "Unique reconciliation keys: " & allKeys.Count, _
        vbInformation, "Reconciliation Universe Test"

End Sub

Sub BuildReconciliationCounts()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    Dim wsRecon As Worksheet
    
    Dim arKeyCol As Long
    Dim bankKeyCol As Long
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    
    Dim arGroups As Object
    Dim bankGroups As Object
    Dim allKeys As Object
    
    Dim i As Long
    Dim outputRow As Long
    Dim key As Variant
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")
    
    arKeyCol = FindHeaderColumn(wsAR, "Invoice_ID")
    bankKeyCol = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arKeyCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankKeyCol).End(xlUp).Row
    
    Set arGroups = CreateObject("Scripting.Dictionary")
    Set bankGroups = CreateObject("Scripting.Dictionary")
    Set allKeys = CreateObject("Scripting.Dictionary")
    
    'Group AR records
    For i = 2 To lastRowAR
        
        key = Trim(CStr(wsAR.Cells(i, arKeyCol).Value))
        
        If key <> "" Then
            
            If arGroups.Exists(key) Then
                arGroups(key) = arGroups(key) + 1
            Else
                arGroups.Add key, 1
            End If
            
            If Not allKeys.Exists(key) Then
                allKeys.Add key, True
            End If
            
        End If
        
    Next i
    
    'Group Bank records
    For i = 2 To lastRowBank
        
        key = Trim(CStr(wsBank.Cells(i, bankKeyCol).Value))
        
        If key <> "" Then
            
            If bankGroups.Exists(key) Then
                bankGroups(key) = bankGroups(key) + 1
            Else
                bankGroups.Add key, 1
            End If
            
            If Not allKeys.Exists(key) Then
                allKeys.Add key, True
            End If
            
        End If
        
    Next i
    
    'Clear previous reconciliation output
    wsRecon.Cells.Clear
    
    'Create headers
    wsRecon.Range("A1").Value = "recon_key"
    wsRecon.Range("B1").Value = "ar_count"
    wsRecon.Range("C1").Value = "bank_count"
    
    outputRow = 2
    
    'Write reconciliation universe
    For Each key In allKeys.Keys
        
        wsRecon.Cells(outputRow, 1).Value = key
        
        If arGroups.Exists(key) Then
            wsRecon.Cells(outputRow, 2).Value = arGroups(key)
        Else
            wsRecon.Cells(outputRow, 2).Value = 0
        End If
        
        If bankGroups.Exists(key) Then
            wsRecon.Cells(outputRow, 3).Value = bankGroups(key)
        Else
            wsRecon.Cells(outputRow, 3).Value = 0
        End If
        
        outputRow = outputRow + 1
        
    Next key
    
    wsRecon.Columns("A:C").AutoFit
    
    MsgBox _
        "Reconciliation count table created." & vbCrLf & vbCrLf & _
        "Reconciliation keys: " & allKeys.Count & vbCrLf & _
        "Output rows: " & outputRow - 2, _
        vbInformation, "Reconciliation Counts"

End Sub


Sub TestFieldComparison()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    Dim wsRecon As Worksheet
    
    Dim arKeyCol As Long
    Dim arCustomerCol As Long
    Dim arAmountCol As Long
    Dim arCurrencyCol As Long
    Dim arDateCol As Long
    
    Dim bankKeyCol As Long
    Dim bankCustomerCol As Long
    Dim bankAmountCol As Long
    Dim bankCurrencyCol As Long
    Dim bankDateCol As Long
    
    Dim lastRowAR As Long
    Dim lastRowBank As Long
    Dim lastRowRecon As Long
    
    Dim arRows As Object
    Dim bankRows As Object
    
    Dim i As Long
    Dim reconRow As Long
    Dim key As String
    
    Dim arRow As Long
    Dim bankRow As Long
    
    Dim amountMatch As Boolean
    Dim currencyMatch As Boolean
    Dim customerMatch As Boolean
    Dim dateMatch As Boolean
    
    Dim dateDifference As Double
    
    Set wsAR = ThisWorkbook.Worksheets("AR_Data")
    Set wsBank = ThisWorkbook.Worksheets("Bank_Data")
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")
    
    'Find AR columns
    arKeyCol = FindHeaderColumn(wsAR, "Invoice_ID")
    arCustomerCol = FindHeaderColumn(wsAR, "Customer_Name")
    arAmountCol = FindHeaderColumn(wsAR, "Invoice_Amount")
    arCurrencyCol = FindHeaderColumn(wsAR, "Currency")
    arDateCol = FindHeaderColumn(wsAR, "Cleared_Date_System")
    
    'Find Bank columns
    bankKeyCol = FindHeaderColumn(wsBank, "Reference_Invoice_ID")
    bankCustomerCol = FindHeaderColumn(wsBank, "Payer_Name")
    bankAmountCol = FindHeaderColumn(wsBank, "Amount_Received")
    bankCurrencyCol = FindHeaderColumn(wsBank, "Currency")
    bankDateCol = FindHeaderColumn(wsBank, "Deposit_Date")
    
    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arKeyCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankKeyCol).End(xlUp).Row
    lastRowRecon = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row
    
    'Create row lookup dictionaries
    Set arRows = CreateObject("Scripting.Dictionary")
    Set bankRows = CreateObject("Scripting.Dictionary")
    
    'Store AR row numbers by key
    For i = 2 To lastRowAR
        
        key = Trim(CStr(wsAR.Cells(i, arKeyCol).Value))
        
        If key <> "" Then
            If Not arRows.Exists(key) Then
                arRows.Add key, i
            End If
        End If
        
    Next i
    
    'Store Bank row numbers by key
    For i = 2 To lastRowBank
        
        key = Trim(CStr(wsBank.Cells(i, bankKeyCol).Value))
        
        If key <> "" Then
            If Not bankRows.Exists(key) Then
                bankRows.Add key, i
            End If
        End If
        
    Next i
    
    'Add comparison headers
    wsRecon.Cells(1, 4).Value = "amount_match"
    wsRecon.Cells(1, 5).Value = "currency_match"
    wsRecon.Cells(1, 6).Value = "customer_match"
    wsRecon.Cells(1, 7).Value = "date_match"
    
    'Compare records
    For reconRow = 2 To lastRowRecon
        
        key = Trim(CStr(wsRecon.Cells(reconRow, 1).Value))
        
        'Only perform field comparison for 1-to-1 records
        If wsRecon.Cells(reconRow, 2).Value = 1 And _
           wsRecon.Cells(reconRow, 3).Value = 1 Then
            
            arRow = arRows(key)
            bankRow = bankRows(key)
            
            'Amount comparison
            amountMatch = _
                Abs(CDbl(wsAR.Cells(arRow, arAmountCol).Value) - _
                    CDbl(wsBank.Cells(bankRow, bankAmountCol).Value)) < 0.01
            
            'Currency comparison
            currencyMatch = _
                UCase(Trim(CStr(wsAR.Cells(arRow, arCurrencyCol).Value))) = _
                UCase(Trim(CStr(wsBank.Cells(bankRow, bankCurrencyCol).Value)))
            
            'Customer comparison
            customerMatch = _
                UCase(Trim(CStr(wsAR.Cells(arRow, arCustomerCol).Value))) = _
                UCase(Trim(CStr(wsBank.Cells(bankRow, bankCustomerCol).Value)))
            
            'Date comparison - +/- 1 day tolerance
            dateDifference = Abs( _
                CDbl(CDate(wsAR.Cells(arRow, arDateCol).Value)) - _
                CDbl(CDate(wsBank.Cells(bankRow, bankDateCol).Value)))
            
            dateMatch = dateDifference <= 1
            
            wsRecon.Cells(reconRow, 4).Value = amountMatch
            wsRecon.Cells(reconRow, 5).Value = currencyMatch
            wsRecon.Cells(reconRow, 6).Value = customerMatch
            wsRecon.Cells(reconRow, 7).Value = dateMatch
            
        Else
            
            wsRecon.Cells(reconRow, 4).Value = "N/A"
            wsRecon.Cells(reconRow, 5).Value = "N/A"
            wsRecon.Cells(reconRow, 6).Value = "N/A"
            wsRecon.Cells(reconRow, 7).Value = "N/A"
            
        End If
        
    Next reconRow
    
    wsRecon.Columns("A:G").AutoFit
    
    MsgBox _
        "Field comparison completed." & vbCrLf & vbCrLf & _
        "Records evaluated: " & lastRowRecon - 1, _
        vbInformation, "Field Comparison"

End Sub


Sub ClassifyReconciliation()

    Dim wsRecon As Worksheet
    
    Dim lastRow As Long
    Dim i As Long
    
    Dim arCount As Long
    Dim bankCount As Long
    
    Dim amountMatch As Variant
    Dim currencyMatch As Variant
    Dim customerMatch As Variant
    Dim dateMatch As Variant
    
    Dim finalStatus As String
    Dim exceptionReason As String
    
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")
    
    lastRow = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row
    
    wsRecon.Cells(1, 8).Value = "final_status"
    wsRecon.Cells(1, 9).Value = "exception_reason"
    
    For i = 2 To lastRow
        
        arCount = wsRecon.Cells(i, 2).Value
        bankCount = wsRecon.Cells(i, 3).Value
        
        amountMatch = wsRecon.Cells(i, 4).Value
        currencyMatch = wsRecon.Cells(i, 5).Value
        customerMatch = wsRecon.Cells(i, 6).Value
        dateMatch = wsRecon.Cells(i, 7).Value
        
        finalStatus = ""
        exceptionReason = ""
        
        'One-sided transactions
        If arCount > 0 And bankCount = 0 Then
            
            finalStatus = "AR_ONLY"
            exceptionReason = "No matching bank transaction"
            
        ElseIf arCount = 0 And bankCount > 0 Then
            
            finalStatus = "BANK_ONLY"
            exceptionReason = "No matching AR transaction"
            
        'Duplicate transactions
        ElseIf arCount > 1 And bankCount > 1 Then
            
            finalStatus = "DUPLICATE_BOTH"
            exceptionReason = "Multiple AR and Bank transactions"
            
        ElseIf arCount > 1 Then
            
            finalStatus = "DUPLICATE_AR"
            exceptionReason = "Multiple AR transactions"
            
        ElseIf bankCount > 1 Then
            
            finalStatus = "DUPLICATE_BANK"
            exceptionReason = "Multiple Bank transactions"
            
        'Normal 1-to-1 reconciliation
        ElseIf arCount = 1 And bankCount = 1 Then
            
            If amountMatch = True And _
               currencyMatch = True And _
               customerMatch = True And _
               dateMatch = True Then
                
                finalStatus = "MATCHED"
                exceptionReason = ""
                
            Else
                
                finalStatus = "EXCEPTION"
                
                If amountMatch = False Then
                    exceptionReason = exceptionReason & "Amount discrepancy; "
                End If
                
                If currencyMatch = False Then
                    exceptionReason = exceptionReason & "Currency discrepancy; "
                End If
                
                If customerMatch = False Then
                    exceptionReason = exceptionReason & "Customer discrepancy; "
                End If
                
                If dateMatch = False Then
                    exceptionReason = exceptionReason & "Date discrepancy; "
                End If
                
            End If
            
        End If
        
        wsRecon.Cells(i, 8).Value = finalStatus
        wsRecon.Cells(i, 9).Value = exceptionReason
        
    Next i
    
    wsRecon.Columns("A:I").AutoFit
    
    MsgBox _
        "Reconciliation classification completed." & vbCrLf & vbCrLf & _
        "Records classified: " & lastRow - 1, _
        vbInformation, "Classification"

End Sub

Sub AddReconciliationActions()

    Dim wsRecon As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim status As String
    Dim action As String
    
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")
    
    lastRow = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row
    
    wsRecon.Cells(1, 10).Value = "action"
    
    For i = 2 To lastRow
        
        status = wsRecon.Cells(i, 8).Value
        
        Select Case status
            
            Case "MATCHED"
                action = "No action required"
                
            Case "AR_ONLY"
                action = "Investigate missing bank transaction"
                
            Case "BANK_ONLY"
                action = "Investigate missing AR transaction"
                
            Case "DUPLICATE_AR"
                action = "Review duplicate AR entries"
                
            Case "DUPLICATE_BANK"
                action = "Review duplicate bank transactions"
                
            Case "DUPLICATE_BOTH"
                action = "Review duplicate transactions on both sides"
                
            Case "EXCEPTION"
                action = "Investigate reconciliation differences"
                
            Case Else
                action = "Review classification"
                
        End Select
        
        wsRecon.Cells(i, 10).Value = action
        
    Next i
    
    wsRecon.Columns("A:J").AutoFit
    
    MsgBox _
        "Reconciliation actions added." & vbCrLf & vbCrLf & _
        "Records updated: " & lastRow - 1, _
        vbInformation, "Reconciliation Actions"

End Sub

Sub ValidateAgainstN8N()

    Dim wsRecon As Worksheet
    Dim wsValidation As Worksheet
    
    Dim lastRow As Long
    Dim i As Long
    
    Dim totalKeys As Long
    Dim matched As Long
    Dim exceptions As Long
    
    Dim arOnly As Long
    Dim bankOnly As Long
    Dim duplicateAR As Long
    Dim duplicateBank As Long
    Dim duplicateBoth As Long
    
    Dim amountDiscrepancy As Long
    Dim currencyDiscrepancy As Long
    Dim customerDiscrepancy As Long
    Dim dateDiscrepancy As Long
    
    Dim status As String
    
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")
    Set wsValidation = ThisWorkbook.Worksheets("Validation")
    
    lastRow = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row
    
    totalKeys = lastRow - 1
    
    For i = 2 To lastRow
        
        status = CStr(wsRecon.Cells(i, 8).Value)
        
        'Overall status
        If status = "MATCHED" Then
            matched = matched + 1
        Else
            exceptions = exceptions + 1
        End If
        
        'Structural exceptions
        Select Case status
        
            Case "AR_ONLY"
                arOnly = arOnly + 1
                
            Case "BANK_ONLY"
                bankOnly = bankOnly + 1
                
            Case "DUPLICATE_AR"
                duplicateAR = duplicateAR + 1
                
            Case "DUPLICATE_BANK"
                duplicateBank = duplicateBank + 1
                
            Case "DUPLICATE_BOTH"
                duplicateBoth = duplicateBoth + 1
                
        End Select
        
        'Field discrepancies
        If wsRecon.Cells(i, 4).Value = False Then
            amountDiscrepancy = amountDiscrepancy + 1
        End If
        
        If wsRecon.Cells(i, 5).Value = False Then
            currencyDiscrepancy = currencyDiscrepancy + 1
        End If
        
        If wsRecon.Cells(i, 6).Value = False Then
            customerDiscrepancy = customerDiscrepancy + 1
        End If
        
        If wsRecon.Cells(i, 7).Value = False Then
            dateDiscrepancy = dateDiscrepancy + 1
        End If
        
    Next i
    
    'Clear old validation results
    wsValidation.Cells.Clear
    
    'Headers
    wsValidation.Range("A1").Value = "Metric"
    wsValidation.Range("B1").Value = "VBA Result"
    wsValidation.Range("C1").Value = "n8n Expected"
    wsValidation.Range("D1").Value = "Difference"
    wsValidation.Range("E1").Value = "Validation"
    
    'Metrics
    wsValidation.Range("A2").Value = "Total reconciliation keys"
    wsValidation.Range("B2").Value = totalKeys
    wsValidation.Range("C2").Value = 1096
    
    wsValidation.Range("A3").Value = "Matched"
    wsValidation.Range("B3").Value = matched
    wsValidation.Range("C3").Value = 365
    
    wsValidation.Range("A4").Value = "Exceptions"
    wsValidation.Range("B4").Value = exceptions
    wsValidation.Range("C4").Value = 731
    
    wsValidation.Range("A5").Value = "AR Only"
    wsValidation.Range("B5").Value = arOnly
    wsValidation.Range("C5").Value = 62
    
    wsValidation.Range("A6").Value = "Bank Only"
    wsValidation.Range("B6").Value = bankOnly
    wsValidation.Range("C6").Value = 43
    
    wsValidation.Range("A7").Value = "Duplicate AR"
    wsValidation.Range("B7").Value = duplicateAR
    wsValidation.Range("C7").Value = 31
    
    wsValidation.Range("A8").Value = "Duplicate Bank"
    wsValidation.Range("B8").Value = duplicateBank
    wsValidation.Range("C8").Value = 38
    
    wsValidation.Range("A9").Value = "Duplicate Both"
    wsValidation.Range("B9").Value = duplicateBoth
    wsValidation.Range("C9").Value = 0
    
    wsValidation.Range("A10").Value = "Amount Discrepancy"
    wsValidation.Range("B10").Value = amountDiscrepancy
    wsValidation.Range("C10").Value = 53
    
    wsValidation.Range("A11").Value = "Currency Discrepancy"
    wsValidation.Range("B11").Value = currencyDiscrepancy
    wsValidation.Range("C11").Value = 59
    
    wsValidation.Range("A12").Value = "Customer Discrepancy"
    wsValidation.Range("B12").Value = customerDiscrepancy
    wsValidation.Range("C12").Value = 67
    
    wsValidation.Range("A13").Value = "Date Discrepancy"
    wsValidation.Range("B13").Value = dateDiscrepancy
    wsValidation.Range("C13").Value = 470
    
    'Calculate differences and PASS/REVIEW
    For i = 2 To 13
        
        wsValidation.Cells(i, 4).Value = _
            wsValidation.Cells(i, 2).Value - wsValidation.Cells(i, 3).Value
        
        If wsValidation.Cells(i, 4).Value = 0 Then
            wsValidation.Cells(i, 5).Value = "PASS"
        Else
            wsValidation.Cells(i, 5).Value = "REVIEW"
        End If
        
    Next i
    
    wsValidation.Columns("A:E").AutoFit
    
    MsgBox _
        "Validation completed." & vbCrLf & vbCrLf & _
        "Check the Validation sheet for VBA vs n8n results.", _
        vbInformation, "Validation Complete"

End Sub

Sub RunReconciliation()

    Application.ScreenUpdating = False

    BuildReconciliationCounts
    TestFieldComparison
    ClassifyReconciliation
    AddReconciliationActions
    ValidateAgainstN8N

    Application.ScreenUpdating = True

    MsgBox _
        "Reconciliation completed successfully." & vbCrLf & vbCrLf & _
        "Check the Reconciliation and Validation sheets.", _
        vbInformation, "Reconciliation Complete"

End Sub

Sub BuildReconciliationCountsPQ()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    Dim wsRecon As Worksheet

    Dim arKeyCol As Long
    Dim bankKeyCol As Long

    Dim lastRowAR As Long
    Dim lastRowBank As Long

    Dim arGroups As Object
    Dim bankGroups As Object
    Dim allKeys As Object

    Dim i As Long
    Dim outputRow As Long
    Dim key As Variant

    Set wsAR = ThisWorkbook.Worksheets("AR_PQ")
    Set wsBank = ThisWorkbook.Worksheets("Bank_PQ")
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")

    arKeyCol = FindHeaderColumn(wsAR, "recon_key")
    bankKeyCol = FindHeaderColumn(wsBank, "recon_key")

    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arKeyCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankKeyCol).End(xlUp).Row

    Set arGroups = CreateObject("Scripting.Dictionary")
    Set bankGroups = CreateObject("Scripting.Dictionary")
    Set allKeys = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRowAR

        key = Trim(CStr(wsAR.Cells(i, arKeyCol).Value))

        If key <> "" Then

            If arGroups.Exists(key) Then
                arGroups(key) = arGroups(key) + 1
            Else
                arGroups.Add key, 1
            End If

            If Not allKeys.Exists(key) Then
                allKeys.Add key, True
            End If

        End If

    Next i

    For i = 2 To lastRowBank

        key = Trim(CStr(wsBank.Cells(i, bankKeyCol).Value))

        If key <> "" Then

            If bankGroups.Exists(key) Then
                bankGroups(key) = bankGroups(key) + 1
            Else
                bankGroups.Add key, 1
            End If

            If Not allKeys.Exists(key) Then
                allKeys.Add key, True
            End If

        End If

    Next i

    wsRecon.Cells.Clear

    wsRecon.Range("A1").Value = "recon_key"
    wsRecon.Range("B1").Value = "ar_count"
    wsRecon.Range("C1").Value = "bank_count"

    outputRow = 2

    For Each key In allKeys.Keys

        wsRecon.Cells(outputRow, 1).Value = key

        If arGroups.Exists(key) Then
            wsRecon.Cells(outputRow, 2).Value = arGroups(key)
        Else
            wsRecon.Cells(outputRow, 2).Value = 0
        End If

        If bankGroups.Exists(key) Then
            wsRecon.Cells(outputRow, 3).Value = bankGroups(key)
        Else
            wsRecon.Cells(outputRow, 3).Value = 0
        End If

        outputRow = outputRow + 1

    Next key

End Sub

Sub ComparePQFields()

    Dim wsAR As Worksheet
    Dim wsBank As Worksheet
    Dim wsRecon As Worksheet

    Dim arKeyCol As Long
    Dim arCustomerCol As Long
    Dim arAmountCol As Long
    Dim arCurrencyCol As Long
    Dim arDateCol As Long

    Dim bankKeyCol As Long
    Dim bankCustomerCol As Long
    Dim bankAmountCol As Long
    Dim bankCurrencyCol As Long
    Dim bankDateCol As Long

    Dim lastRowAR As Long
    Dim lastRowBank As Long
    Dim lastRowRecon As Long

    Dim arRows As Object
    Dim bankRows As Object

    Dim i As Long
    Dim reconRow As Long
    Dim key As String

    Dim arRow As Long
    Dim bankRow As Long

    Dim amountMatch As Boolean
    Dim currencyMatch As Boolean
    Dim customerMatch As Boolean
    Dim dateMatch As Boolean

    Dim dateDifference As Double

    Set wsAR = ThisWorkbook.Worksheets("AR_PQ")
    Set wsBank = ThisWorkbook.Worksheets("Bank_PQ")
    Set wsRecon = ThisWorkbook.Worksheets("Reconciliation")

    arKeyCol = FindHeaderColumn(wsAR, "recon_key")
    arCustomerCol = FindHeaderColumn(wsAR, "customer")
    arAmountCol = FindHeaderColumn(wsAR, "amount")
    arCurrencyCol = FindHeaderColumn(wsAR, "currency")
    arDateCol = FindHeaderColumn(wsAR, "transaction_date")

    bankKeyCol = FindHeaderColumn(wsBank, "recon_key")
    bankCustomerCol = FindHeaderColumn(wsBank, "customer")
    bankAmountCol = FindHeaderColumn(wsBank, "amount")
    bankCurrencyCol = FindHeaderColumn(wsBank, "currency")
    bankDateCol = FindHeaderColumn(wsBank, "transaction_date")

    lastRowAR = wsAR.Cells(wsAR.Rows.Count, arKeyCol).End(xlUp).Row
    lastRowBank = wsBank.Cells(wsBank.Rows.Count, bankKeyCol).End(xlUp).Row
    lastRowRecon = wsRecon.Cells(wsRecon.Rows.Count, 1).End(xlUp).Row

    Set arRows = CreateObject("Scripting.Dictionary")
    Set bankRows = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRowAR

        key = Trim(CStr(wsAR.Cells(i, arKeyCol).Value))

        If key <> "" Then
            If Not arRows.Exists(key) Then
                arRows.Add key, i
            End If
        End If

    Next i

    For i = 2 To lastRowBank

        key = Trim(CStr(wsBank.Cells(i, bankKeyCol).Value))

        If key <> "" Then
            If Not bankRows.Exists(key) Then
                bankRows.Add key, i
            End If
        End If

    Next i

    wsRecon.Cells(1, 4).Value = "amount_match"
    wsRecon.Cells(1, 5).Value = "currency_match"
    wsRecon.Cells(1, 6).Value = "customer_match"
    wsRecon.Cells(1, 7).Value = "date_match"

    For reconRow = 2 To lastRowRecon

        key = Trim(CStr(wsRecon.Cells(reconRow, 1).Value))

        If wsRecon.Cells(reconRow, 2).Value = 1 And _
           wsRecon.Cells(reconRow, 3).Value = 1 Then

            arRow = arRows(key)
            bankRow = bankRows(key)

            amountMatch = _
                Abs(CDbl(wsAR.Cells(arRow, arAmountCol).Value) - _
                    CDbl(wsBank.Cells(bankRow, bankAmountCol).Value)) < 0.01

            currencyMatch = _
                UCase(Trim(CStr(wsAR.Cells(arRow, arCurrencyCol).Value))) = _
                UCase(Trim(CStr(wsBank.Cells(bankRow, bankCurrencyCol).Value)))

            customerMatch = _
                UCase(Trim(CStr(wsAR.Cells(arRow, arCustomerCol).Value))) = _
                UCase(Trim(CStr(wsBank.Cells(bankRow, bankCustomerCol).Value)))

            dateDifference = Abs( _
                CDbl(CDate(wsAR.Cells(arRow, arDateCol).Value)) - _
                CDbl(CDate(wsBank.Cells(bankRow, bankDateCol).Value)))

            dateMatch = dateDifference <= 1

            wsRecon.Cells(reconRow, 4).Value = amountMatch
            wsRecon.Cells(reconRow, 5).Value = currencyMatch
            wsRecon.Cells(reconRow, 6).Value = customerMatch
            wsRecon.Cells(reconRow, 7).Value = dateMatch

        Else

            wsRecon.Cells(reconRow, 4).Value = "N/A"
            wsRecon.Cells(reconRow, 5).Value = "N/A"
            wsRecon.Cells(reconRow, 6).Value = "N/A"
            wsRecon.Cells(reconRow, 7).Value = "N/A"

        End If

    Next reconRow

End Sub

Sub RunPQReconciliation()

    Application.ScreenUpdating = False

    BuildReconciliationCountsPQ
    ComparePQFields
    ClassifyReconciliation
    AddReconciliationActions
    ValidateAgainstN8N

    Application.ScreenUpdating = True

    MsgBox _
        "Power Query + VBA reconciliation completed." & vbCrLf & vbCrLf & _
        "Check the Validation sheet.", _
        vbInformation, "Automation Complete"

End Sub


