Imports System.Data.OleDb

Module ace
    Private connection As New OleDbConnection
    Private table() As String
    Private da As New OleDbDataAdapter

    Private Sub gettable()
        Dim schemaTable As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        Dim index As Integer = 0
        For i = 0 To schemaTable.Rows.Count - 1
            If InStr(schemaTable.Rows(i).Item("TABLE_NAME"), "~") = 0 Then
                ReDim Preserve table(index)
                table(index) = schemaTable.Rows(i).Item("TABLE_NAME")
                index = index + 1
            End If
        Next
    End Sub

    Public Sub load(ByVal dbpath As String, ByRef ds As DataSet)
        connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbpath
        connection.Open()
        gettable()
        ds.Clear()
        If table.Length > 0 Then
            For i = 0 To table.Length - 1
                da.SelectCommand = New OleDbCommand("SELECT * FROM [" & table(i) & "]", connection)
                Try
                    da.Fill(ds, table(i))
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString)
                End Try

            Next

        End If
        connection.Close()
    End Sub

    Public Sub save(ByVal ds As DataSet, ByVal dbpath As String)
        connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbpath
        connection.Open()
        gettable()
        For i = 0 To table.Length - 1
            da.SelectCommand = New OleDbCommand("SELECT * FROM " & table(i), connection)
            Dim objCommandBuilder As New OleDbCommandBuilder(da)
            objCommandBuilder.QuotePrefix = "[" : objCommandBuilder.QuoteSuffix = "]"
            Try
                da.Update(ds, table(i))
            Catch ex As Exception
            End Try
        Next
        connection.Close()
    End Sub

    Public Function add(ByRef dt As DataTable) As DataRow
        Dim MaxID As Integer = 0
        If Not String.IsNullOrEmpty(dt.Compute("max(ID)", String.Empty).ToString) Then MaxID = Convert.ToInt32(dt.Compute("max(ID)", String.Empty))
        Dim newrow As DataRow = Nothing
        newrow = dt.NewRow()
        MaxID = MaxID + 1
        newrow("ID") = MaxID.ToString

        Try
            dt.Rows.Add(newrow)
        Catch ex As Exception

        End Try

        add = newrow
    End Function

    Public Sub remove(ByRef dt As DataTable, ByRef row As DataRow)
        row.BeginEdit()
        row.Delete()
        row.EndEdit()
    End Sub

End Module
