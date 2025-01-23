Option Explicit

Public Sub ImportDataFromAccess()
    ' Declare variables
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim dbPath As String
    Dim sqlQuery As String
    Dim lastRow As Long
    Dim i As Integer

    
    ' Set the path to the Access database
    dbPath = ThisWorkbook.Path & "\project.accdb"
    
    ' Check if the database file exists
    If Dir(dbPath) = "" Then
        MsgBox "Database file not found at: " & dbPath, vbExclamation
        Exit Sub
    End If
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("Database")
    ws.Cells.Clear ' Clear the sheet
    
    ' Create ADO connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    ' Construct SQL query to join the three tables
    sqlQuery = "SELECT TOP 1000 " & _
               "Orders.order_id, Orders.customer_id, Orders.payment_value, Orders.payment_type, Orders.order_status, Orders.order_purchase_timestamp, " & _
               "Products.product_id, Products.product_category_name_english, Products.price, " & _
               "Customers.customer_city, Customers.customer_state " & _
               "FROM (Orders INNER JOIN Products ON Orders.order_id = Products.order_id) " & _
               "INNER JOIN Customers ON Orders.customer_id = Customers.customer_id;"
    
    ' Execute the query
    Set rs = conn.Execute(sqlQuery)
    
    ' Write field names in the first row
    For i = 1 To rs.Fields.Count
        ws.Cells(1, i).Value = rs.Fields(i - 1).Name
    Next i
    
    ' Write data to the worksheet
    lastRow = 2
    Do Until rs.EOF
        For i = 1 To rs.Fields.Count
            ws.Cells(lastRow, i).Value = rs.Fields(i - 1).Value
        Next i
        rs.MoveNext
        lastRow = lastRow + 1
    Loop
    
    ' Clean up objects
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    ' Automatically jump to the "Database" sheet for user visibility
    ws.Activate
    
    MsgBox "Data successfully imported into the Database sheet!", vbInformation

End Sub

Public Sub FetchOrdersBySelectedStatus()
    ' Declaring database and recordset objects
    Dim dbConnection As ADODB.Connection
    Dim orderRecords As ADODB.Recordset
    Dim sqlQuery As String
    Dim selectedStatus As String
    Dim nextRow As Long
    Dim wsHome As Worksheet
    Dim wsOrders As Worksheet
    
    ' Relative path to database
    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\project.accdb"

    ' Check if the database file exists
    If Dir(dbPath) = "" Then
        MsgBox "Database file not found at: " & dbPath, vbExclamation
        Exit Sub
    End If

    ' Reference to the Home page and Orders worksheets
    Set wsHome = ThisWorkbook.Sheets("Home page")
    Set wsOrders = ThisWorkbook.Sheets("Orders")

    ' Read the delivery status from a named range "DeliveryStatus" on the Home page
    selectedStatus = wsHome.Range("DeliveryStatus").Value

    ' Validate the input for delivery status
    If selectedStatus = "" Then
        MsgBox "Please enter a valid order status in the named range 'DeliveryStatus' on the Home page."
        Exit Sub
    End If

    ' Clear previous data and formats in the Orders worksheet
    wsOrders.Range("A2:F1048576").ClearContents
    wsOrders.Range("A2:F1048576").ClearFormats

    ' Initialize the connection to the database
    Set dbConnection = New ADODB.Connection
    dbConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' SQL query to retrieve orders based on the selected status
    sqlQuery = "SELECT order_id, customer_id, payment_value, payment_type, order_status, order_purchase_timestamp " & _
               "FROM Orders WHERE order_status = '" & selectedStatus & "'"

    ' Open the recordset
    Set orderRecords = New ADODB.Recordset
    orderRecords.Open sqlQuery, dbConnection, adOpenStatic, adLockReadOnly

    ' Headers setup in the Orders sheet
    With wsOrders
        .Cells(1, 1).Value = "Order ID"
        .Cells(1, 2).Value = "Customer ID"
        .Cells(1, 3).Value = "Payment Value"
        .Cells(1, 4).Value = "Payment Type"
        .Cells(1, 5).Value = "Delivery Status"
        .Cells(1, 6).Value = "Purchase Timestamp"

        ' Populate data starting from the second row
        nextRow = 2
        Do While Not orderRecords.EOF
            .Cells(nextRow, 1).Value = orderRecords("order_id")
            .Cells(nextRow, 2).Value = orderRecords("customer_id")
            .Cells(nextRow, 3).Value = orderRecords("payment_value")
            .Cells(nextRow, 4).Value = orderRecords("payment_type")
            .Cells(nextRow, 5).Value = orderRecords("order_status")
            .Cells(nextRow, 6).Value = orderRecords("order_purchase_timestamp")

            ' Apply color coding based on the delivery status
            Select Case .Cells(nextRow, 5).Value
                Case "pending"
                    .Cells(nextRow, 5).Interior.Color = RGB(255, 223, 186)  ' Light Orange
                Case "shipped"
                    .Cells(nextRow, 5).Interior.Color = RGB(155, 194, 230)  ' Light Blue
                Case "delivered"
                    .Cells(nextRow, 5).Interior.Color = RGB(169, 208, 142)  ' Light Green
            End Select

            nextRow = nextRow + 1
            orderRecords.MoveNext
        Loop
    End With

    ' Close the database connection and clean up objects
    orderRecords.Close
    Set orderRecords = Nothing
    dbConnection.Close
    Set dbConnection = Nothing
    ' Jump to the "Orders" sheet
    wsOrders.Activate
    
End Sub

Public Sub DisplayOrderStatusCounts()
    ' Declaring database connection and recordset variables
    Dim dbConnection As ADODB.Connection
    Dim orderStatusRecordset As ADODB.Recordset
    Dim sqlQuery As String
    Dim summaryWorksheet As Worksheet
    
    ' Relative path to database
    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\project.accdb"

    ' Check if the database file exists
    If Dir(dbPath) = "" Then
        MsgBox "Database file not found at: " & dbPath, vbExclamation
        Exit Sub
    End If
    
    ' Setting references to the worksheet
    Set summaryWorksheet = ThisWorkbook.Sheets("orders_status_distribution")
    summaryWorksheet.Cells.ClearContents ' Clear previous contents

    ' Establishing connection to the database
    Set dbConnection = New ADODB.Connection
    dbConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' SQL query to fetch the count of orders by their status and order them in descending order by count
    sqlQuery = "SELECT order_status, COUNT(order_id) AS OrderCount FROM Orders GROUP BY order_status " & _
               "ORDER BY COUNT(order_id) DESC"

    ' Execute the query and set the recordset
    Set orderStatusRecordset = New ADODB.Recordset
    orderStatusRecordset.Open sqlQuery, dbConnection, adOpenStatic, adLockReadOnly

    ' Setup headers for data display
    With summaryWorksheet
        .Cells(1, 1).Value = "Order Status"
        .Cells(1, 2).Value = "Number of Orders"
        Dim rowIndex As Integer
        rowIndex = 2
        Do While Not orderStatusRecordset.EOF
            .Cells(rowIndex, 1).Value = orderStatusRecordset("order_status")
            .Cells(rowIndex, 2).Value = orderStatusRecordset("OrderCount")
            rowIndex = rowIndex + 1
            orderStatusRecordset.MoveNext
        Loop
    End With

    ' Cleanup database connections
    orderStatusRecordset.Close
    Set orderStatusRecordset = Nothing
    dbConnection.Close
    Set dbConnection = Nothing
    summaryWorksheet.Activate
End Sub

Public Sub DisplayCustomerOrders()
    ' Declaring database connection, command, and recordset variables
    Dim dbConnection As ADODB.Connection
    Dim customerOrders As ADODB.Recordset
    Dim customerInfo As ADODB.Recordset
    Dim sqlQuery As String
    Dim customerID As String
    Dim homeSheet As Worksheet, returnSheet As Worksheet

    ' Relative path to database
    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\project.accdb"

    ' Check if the database file exists
    If Dir(dbPath) = "" Then
        MsgBox "Database file not found at: " & dbPath, vbExclamation
        Exit Sub
    End If
    
    ' Setting references to the worksheets
    Set homeSheet = ThisWorkbook.Sheets("Home page")
    Set returnSheet = ThisWorkbook.Sheets("customer_id_return")
    returnSheet.Cells.ClearContents ' Clear previous contents

    ' Retrieve customer ID from the Home page
    customerID = homeSheet.Range("Customer_id").Value
    If customerID = "" Then
        MsgBox "Please enter a customer ID in the designated cell."
        Exit Sub
    End If

    ' Establishing connection to the database
    Set dbConnection = New ADODB.Connection
    dbConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' SQL query to fetch the customer's order details
    sqlQuery = "SELECT o.order_id, o.customer_id, o.payment_value, o.payment_type, o.order_status, o.order_purchase_timestamp, " & _
               "c.customer_city, c.customer_state FROM Orders o INNER JOIN Customers c ON o.customer_id = c.customer_id " & _
               "WHERE o.customer_id = '" & customerID & "'"

    ' Execute the query and set the recordset
    Set customerOrders = New ADODB.Recordset
    customerOrders.Open sqlQuery, dbConnection, adOpenStatic, adLockReadOnly

    ' Setup headers and start populating data
    With returnSheet
        .Cells(1, 1).Value = "Order ID"
        .Cells(1, 2).Value = "Payment Value"
        .Cells(1, 3).Value = "Payment Type"
        .Cells(1, 4).Value = "Order Status"
        .Cells(1, 5).Value = "Purchase Date"
        .Cells(1, 6).Value = "Customer ID"
        .Cells(1, 7).Value = "Customer City"
        .Cells(1, 8).Value = "Customer State"
        
        ' Set the format for the date column
        .Columns(5).NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
    
        Dim rowIndex As Integer
        rowIndex = 2
        Do While Not customerOrders.EOF
            .Cells(rowIndex, 1).Value = customerOrders("order_id")
            .Cells(rowIndex, 2).Value = customerOrders("payment_value")
            .Cells(rowIndex, 3).Value = customerOrders("payment_type")
            .Cells(rowIndex, 4).Value = customerOrders("order_status")
            .Cells(rowIndex, 5).Value = customerOrders("order_purchase_timestamp")
            .Cells(rowIndex, 6).Value = customerOrders("customer_id")
            .Cells(rowIndex, 7).Value = customerOrders("customer_city")
            .Cells(rowIndex, 8).Value = customerOrders("customer_state")
            rowIndex = rowIndex + 1
            customerOrders.MoveNext
        Loop
    End With

    ' Cleanup database connections
    customerOrders.Close
    Set customerOrders = Nothing
    dbConnection.Close
    Set dbConnection = Nothing
    
    returnSheet.Activate
  
End Sub


Sub InsertDataIntoAccess()
    ' Declare variables
    Dim conn As Object
    Dim ws As Worksheet
    Dim dbPath As String
    Dim lastRow As Long
    Dim sqlQuery As String
    Dim i As Long
    
    ' Set the path to the Access database
    dbPath = ThisWorkbook.Path & "\project.accdb"
    
    ' Set the worksheet containing new records
    Set ws = ThisWorkbook.Sheets("new_records")
    
    ' Determine the last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Assuming data starts in column A
    
    ' Create ADO connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    ' Loop through each row in new_records
    For i = 2 To lastRow ' Assuming row 1 has headers
        ' Construct the SQL query to insert data
        sqlQuery = "INSERT INTO Orders (order_id, customer_id, payment_value, payment_type, order_status, order_purchase_timestamp) " & _
                   "VALUES ('" & ws.Cells(i, 1).Value & "', '" & ws.Cells(i, 2).Value & "', " & ws.Cells(i, 3).Value & ", '" & _
                   ws.Cells(i, 4).Value & "', '" & ws.Cells(i, 5).Value & "', #" & ws.Cells(i, 6).Value & "#);"
        
        ' Execute the query
        On Error Resume Next ' Handle errors gracefully
        conn.Execute sqlQuery
        If Err.Number <> 0 Then
            MsgBox "Error inserting row " & i & ": " & Err.Description, vbExclamation
            Err.Clear
        End If
    Next i
    
    ' Clean up
    conn.Close
    Set conn = Nothing
    
    ws.Activate
    
    MsgBox "Data successfully inserted into the Access database!", vbInformation
End Sub