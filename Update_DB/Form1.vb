Imports System.Data.OleDb
Imports System.Text

Public Class Form1
    Dim sourceConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\updata.accdb;"
    Dim destinationConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\Online Data - PROLAB.SQL;"
    Dim value
    Dim schemaTables As DataTable
    Sub table_get(ByVal ins)
        Dim connectionString As String = sourceConnectionString


        '  Dim connectionString As String = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath}"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            schemaTables = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tableNames, Nothing})

            ' Get the DataTable that contains information about the tables in the database.
            Dim schemaTable As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

            If schemaTable IsNot Nothing Then
                ' Iterate through the rows in the DataTable to get the table names.
                For Each row As DataRow In schemaTable.Rows
                    ' The table name is in the "TABLE_NAME" column.
                    Dim tableName As String = row("TABLE_NAME").ToString()
                    ' You might want to filter out system tables like "MSys...".
                    If Not tableName.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) Then
                        If ins = 1 Then
                            CopyTableButton(tableName)
                        Else

                            value = "---------" & tableName & "----------"
                            If value IsNot Nothing Then
                                ' قم بتحديث النص في TextBox باستخدام القيمة الممررة
                                If Me.InvokeRequired Then
                                    Me.Invoke(New Action(Of String)(AddressOf UpdateTextOnUIThread2), value)
                                Else
                                    UpdateTextOnUIThread2(value)
                                End If
                            End If
                            If value IsNot Nothing Then
                                ' قم بتحديث النص في TextBox باستخدام القيمة الممررة
                                If Me.InvokeRequired Then
                                    Me.Invoke(New Action(Of String)(AddressOf UpdateTextOnUIThread), value)
                                Else
                                    UpdateTextOnUIThread(value)
                                End If
                            End If
                            updatetaple(tableName)
                        End If
                    End If
                Next
            End If

            connection.Close()
        End Using
    End Sub
    Private Sub CopyTableButton(ByVal tableName As String)
        On Error Resume Next
        ' Use the source and destination database paths defined in Button1_Click.Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\updata.SQL;

        Using sourceConnection As New OleDbConnection(sourceConnectionString)
            Using destinationConnection As New OleDbConnection(destinationConnectionString)
                sourceConnection.Open()
                destinationConnection.Open()

                ' Define the SQL command to select data from the source table.
                Dim selectCommand As New OleDbCommand($"SELECT * FROM [{tableName}]", sourceConnection)

                ' Create a data adapter to fetch data from the source table.
                Dim dataAdapter As New OleDbDataAdapter(selectCommand)

                ' Create a DataTable to hold the table data.
                Dim dataTable As New DataTable()

                ' Fill the DataTable with data from the source table.
                dataAdapter.Fill(dataTable)

                ' Create a new table in the destination database.
                Dim createTableCommand As New OleDbCommand($"CREATE TABLE [{tableName}] ", destinationConnection)

                createTableCommand.ExecuteNonQuery()

                ' Use a data adapter to update the destination table with the data from the DataTable.
                Using dataAdapterDestination As New OleDbDataAdapter()
                    dataAdapterDestination.SelectCommand = selectCommand
                    Dim commandBuilder As New OleDbCommandBuilder(dataAdapterDestination)
                    dataAdapterDestination.InsertCommand = commandBuilder.GetInsertCommand()
                    dataAdapterDestination.Update(dataTable)
                End Using



                ' Close connections.
                sourceConnection.Close()
                destinationConnection.Close()
            End Using
        End Using

    End Sub

    Private Function GetTableColumns(dataTable As DataTable) As String
        ' On Error Resume Next
        ' Build a string representing the columns and their data types.

        Dim columnString As New StringBuilder()
        Dim isAutoIncrements As Boolean = False


        For Each column As DataColumn In dataTable.Columns

            isAutoIncrement = IsAutoIncrementColumn(schemaTable, tableName, column.ColumnName)


            columnString.Append($"[{column.ColumnName}] {GetAccessDataType(column.DataType, isAutoIncrements)}, ")

        Next

        If columnString.Length > 2 Then
            columnString.Length -= 2
        End If

        Return columnString.ToString()


    End Function
    Private Function IsAutoIncrementColumn(schemaTable As DataTable, tableName As String, columnName As String) As Boolean
        ' Iterate through the schema table to find information about the specified column
        For Each row As DataRow In schemaTable.Rows
            If String.Equals(row("TABLE_NAME").ToString(), tableName, StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(row("COLUMN_NAME").ToString(), columnName, StringComparison.OrdinalIgnoreCase) Then

                ' Check the COLUMN_FLAGS to determine if it's an auto-increment column
                Dim columnFlags As Integer = CInt(row("COLUMN_FLAGS"))
                Return (columnFlags And &H10) = &H10  ' Check the auto-increment flag (bit 4)
            End If
        Next

        ' Column not found or not auto-increment
        Return False
    End Function
    Function GetAccessDataType(dataType As Type, isAutoIncrementd As Boolean) As String
        ' Map .NET data types to Access data types.
        Select Case Type.GetTypeCode(dataType)
            Case TypeCode.String
                Return "TEXT"
            Case TypeCode.Int32
                If isAutoIncrementd Then
                    Return "COUNTER"
                Else
                    Return "LONG"
                End If
            Case TypeCode.Double
                Return "DOUBLE"
            Case TypeCode.Decimal
                Return "DECIMAL"
            Case TypeCode.Boolean
                Return "YESNO"
            Case TypeCode.DateTime
                Return "DATETIME"
            Case Else
                Return "TEXT"
        End Select
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        table_get(1)

        MessageBox.Show(" successfully. insert First Step ")
        table_get(2)
        MessageBox.Show(" successfully. Done !")
    End Sub

    Sub updatetaple(ByVal tableName)
        ' On Error Resume Next
        ' Connection strings for the source and destination databases


        ' Open connections
        Using sourceConnection As New OleDbConnection(sourceConnectionString),
              destinationConnection As New OleDbConnection(destinationConnectionString)
            Try
                ' Open the connections
                sourceConnection.Open()
                destinationConnection.Open()
            Catch ex As Exception
                value = ex.Message
                If value IsNot Nothing Then
                    ' قم بتحديث النص في TextBox باستخدام القيمة الممررة
                    If Me.InvokeRequired Then
                        Me.Invoke(New Action(Of String)(AddressOf UpdateTextOnUIThread), value)
                    Else
                        UpdateTextOnUIThread(value)
                    End If
                End If
            End Try


            ' Retrieve schema information for the "Web_visit" table from the source database
            Dim schemaTable As DataTable = sourceConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                New Object() {Nothing, Nothing, $"{tableName}", Nothing})

            ' Iterate through the columns and create ALTER TABLE statements
            For Each row As DataRow In schemaTable.Rows
                Dim columnName As String = row("COLUMN_NAME").ToString()
                Dim dataType = row("DATA_TYPE").ToString
                Dim size As Integer = If(row("CHARACTER_MAXIMUM_LENGTH") IsNot DBNull.Value, CInt(row("CHARACTER_MAXIMUM_LENGTH")), -1)
                Dim columnFlags As Integer = CInt(row("COLUMN_FLAGS"))

                ' Check if the column is auto-increment
                Dim isAutoIncrement As Boolean = (columnFlags And &H10) = &H10

                ' Check if the column is a primary key
                Dim isPrimaryKey As Boolean = (columnFlags And &H2) = &H2


                ' Create ALTER TABLE statement based on the column information
                Dim alterTableSql As String = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {GetAccessDataTypes(dataType, size, isAutoIncrement)};"
                Try
                    ' Execute the ALTER TABLE statement on the destination database
                    Using command As New OleDbCommand(alterTableSql, destinationConnection)
                        command.ExecuteNonQuery()
                    End Using


                    value = tableName & $"({dataType.ToString})" & "->" & columnName & "->" & GetAccessDataTypes(dataType, size, isAutoIncrement) & $",{isAutoIncrement}"
                    If value IsNot Nothing Then
                        ' قم بتحديث النص في TextBox باستخدام القيمة الممررة
                        If Me.InvokeRequired Then
                            Me.Invoke(New Action(Of String)(AddressOf UpdateTextOnUIThread2), value)
                        Else
                            UpdateTextOnUIThread2(value)
                        End If
                    End If
                Catch ex As Exception

                    value = ex.Message
                    If value IsNot Nothing Then
                        ' قم بتحديث النص في TextBox باستخدام القيمة الممررة
                        If Me.InvokeRequired Then
                            Me.Invoke(New Action(Of String)(AddressOf UpdateTextOnUIThread), value)
                        Else
                            UpdateTextOnUIThread(value)
                        End If
                    End If
                End Try

            Next
        End Using
    End Sub
    Function GetAccessDataTypes(ByVal dataType As Integer, ByVal size As Integer, ByVal isAutoIncrement As Boolean) As String


        Select Case dataType
            Case 11 ' OleDbType.Boolean
                Return "YESNO"
            Case 3 ' OleDbType.Integer
                If isAutoIncrement Then
                    Return "COUNTER PRIMARY KEY"
                Else
                    Return "LONG"
                End If
            Case 130 ' OleDbType.String
                Return If(size = -1, "MEMO", $"VARCHAR({size})")
            Case 4 ' OleDbType.Double
                Return "DOUBLE"
            Case 5 ' OleDbType.Currency
                Return "CURRENCY"
            Case 7 ' OleDbType.DateTime
                Return "DATETIME"
                ' Add more cases as needed for other data types
            Case Else
                Return "TEXT"
        End Select
    End Function

    Private Sub UpdateTextOnUIThread(value As String)
        ' Update the TextBox text here
        RichTextBox1.AppendText(vbCrLf & value)
        ' Set the SelectionStart to the end of the text
        RichTextBox1.SelectionStart = RichTextBox1.Text.Length

        ' Scroll to the end to show the last text added
        RichTextBox1.ScrollToCaret()
    End Sub
    Private Sub UpdateTextOnUIThread2(value As String)
        ' Update the TextBox text here
        RichTextBox2.AppendText(vbCrLf & value)
        ' Set the SelectionStart to the end of the text
        RichTextBox2.SelectionStart = RichTextBox2.Text.Length

        ' Scroll to the end to show the last text added
        RichTextBox2.ScrollToCaret()
    End Sub
    ' Function to map Access data types to equivalent VB.NET data types



End Class
