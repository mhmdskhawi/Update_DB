Imports System.Data.OleDb
Imports System.Text

Public Class Form1
    Dim sourceConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\updata.SQL;"
    Dim destinationConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\Online Data - PROLAB.SQL;"


    Sub table_get(ByVal ins)
        Dim connectionString As String = sourceConnectionString


        '  Dim connectionString As String = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={databasePath}"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()

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
                Dim createTableCommand As New OleDbCommand($"CREATE TABLE [{tableName}] ({GetTableColumns(dataTable)})", destinationConnection)

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
        ' Build a string representing the columns and their data types.
        Dim columnString As New StringBuilder()

        For Each column As DataColumn In dataTable.Columns
            columnString.Append($"[{column.ColumnName}] {GetAccessDataType(column.DataType)}, ")
        Next

        ' Remove the trailing comma and space.
        columnString.Length -= 2

        Return columnString.ToString()
    End Function

    Function GetAccessDataType(dataType As Type) As String
        ' Map .NET data types to Access data types.
        Select Case Type.GetTypeCode(dataType)
            Case TypeCode.String
                Return "TEXT"
            Case TypeCode.Int32
                Return "LONG"
            Case TypeCode.Double
                Return "DOUBLE"
            Case TypeCode.Decimal
                Return "DECIMAL"
            Case TypeCode.Int32
                Return "INTEGER"
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
        On Error Resume Next
        ' Connection strings for the source and destination databases


        ' Open connections
        Using sourceConnection As New OleDbConnection(sourceConnectionString),
              destinationConnection As New OleDbConnection(destinationConnectionString)

            ' Open the connections
            sourceConnection.Open()
            destinationConnection.Open()

            ' Retrieve schema information for the "Web_visit" table from the source database
            Dim schemaTable As DataTable = sourceConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                New Object() {Nothing, Nothing, $"{tableName}", Nothing})

            ' Iterate through the columns and create ALTER TABLE statements
            For Each row As DataRow In schemaTable.Rows
                Dim columnName As String = row("COLUMN_NAME").ToString()
                Dim dataType As Type = row("DATA_TYPE").GetType
                Dim size As Integer = If(row("CHARACTER_MAXIMUM_LENGTH") IsNot DBNull.Value, CInt(row("CHARACTER_MAXIMUM_LENGTH")), -1)

                ' Create ALTER TABLE statement based on the column information
                Dim alterTableSql As String = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {GetAccessDataType(dataType)};"

                ' Execute the ALTER TABLE statement on the destination database
                Using command As New OleDbCommand(alterTableSql, destinationConnection)
                    command.ExecuteNonQuery()
                End Using
            Next
        End Using
    End Sub

    ' Function to map Access data types to equivalent VB.NET data types



End Class
