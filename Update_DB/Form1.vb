Imports System.Data.OleDb
Imports System.Text

Public Class Form1


    Sub table_get()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\updata.SQL;"


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

                        CopyTableButton(tableName)
                    End If
                Next
            End If

            connection.Close()
        End Using
    End Sub
    Private Sub CopyTableButton(ByVal tableName As String)
        On Error Resume Next
        ' Use the source and destination database paths defined in Button1_Click.Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\TECH LAB SYSTEM\Data\updata.SQL;
        Dim sourceDatabasePath As String = "D:\TECH LAB SYSTEM\Data\updata.SQL"
        Dim destinationDatabasePath As String = "D:\TECH LAB SYSTEM\Data\Online Data - PROLAB.SQL"

        Dim connectionStringSource As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={sourceDatabasePath}"
        Dim connectionStringDestination As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={destinationDatabasePath}"

        Using sourceConnection As New OleDbConnection(connectionStringSource)
            Using destinationConnection As New OleDbConnection(connectionStringDestination)
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

                '  MessageBox.Show($"Table '{tableName}' copied successfully.")

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

    Private Function GetAccessDataType(dataType As Type) As String
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
            Case TypeCode.DateTime
                Return "DATETIME"
            Case Else
                Return "TEXT"
        End Select
    End Function




End Class
