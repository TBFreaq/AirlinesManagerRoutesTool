Imports System.Data.OleDb

Public Class AccDBHandler

    Dim Source As String = ""
    Dim FileName As String = ""
    Dim CompletePath As String = ""
    Dim ConnectionOpen As Boolean = False
    Public DataSetCollection As New DataSet
    Dim DatabaseConnection As New OleDbConnection
    Dim MDBConnectionString As String
    Dim DataReader As OleDbDataReader

    Public Sub SetSource(SourceStr As String, FileNameStr As String)
        Source = SourceStr
        FileName = FileNameStr
        CompletePath = Source & FileName
    End Sub

    Public Sub CreateNewDB(TableName As String, Columns As String(), Optional Overwrite As Boolean = False)

        If CompletePath = "" Then
            MsgBox("No Source set")
            Exit Sub
        End If

        If Overwrite = True Then

            If IO.File.Exists(CompletePath) Then
                IO.File.Delete(CompletePath)
            End If

        Else

            If IO.File.Exists(CompletePath) Then
                MsgBox("DB already exists")
                Exit Sub
            End If

        End If


        If Not IO.Directory.Exists(Source) Then
            IO.Directory.CreateDirectory(Source)
        End If

        Dim ADOXCatalog As New ADOX.Catalog
        Dim ADOXTable As New ADOX.Table
        Dim ADOXTable2 As New ADOX.Table
        Dim ADOXIndex As New ADOX.Index

        ADOXCatalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CompletePath)

        On Error Resume Next

        ADOXCatalog.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CompletePath

        ADOXTable.Name = TableName
        For Each ColumnName In Columns
            ADOXTable.Columns.Append(ColumnName, ADOX.DataTypeEnum.adLongVarWChar)
        Next
        ADOXCatalog.Tables.Append(ADOXTable)
        ADOXTable.Indexes.Append(ADOXIndex)

        ADOXTable = Nothing
        ADOXTable2 = Nothing
        ADOXCatalog = Nothing
        ADOXIndex = Nothing

    End Sub

    Public Sub ReadDatabase(Table As String, Optional ReadAll As Boolean = True, Optional SearchString As String = "")
        Dim Query As String
        MDBConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CompletePath

        If ConnectionOpen = False Then
            DatabaseConnection.ConnectionString = MDBConnectionString
            Try
                DatabaseConnection.Open()
                ConnectionOpen = True
            Catch ex As Exception
                MsgBox(ex)
                Exit Sub
            End Try
        End If

        Query = "SELECT * FROM " & Table

        Dim DatabaseCommand As New OleDbCommand(Query, DatabaseConnection)
        Dim DatabaseAdapter As New OleDbDataAdapter(DatabaseCommand)

        DataReader = DatabaseCommand.ExecuteReader

        Form1.LVRoutes.Items.Clear()

        If DataReader.HasRows Then

            While DataReader.Read
                Dim NewListViewItem As New ListViewItem
                NewListViewItem.Text = DataReader.GetValue(0) 'Route Name
                NewListViewItem.SubItems.Add(DataReader.GetValue(1)) 'Route Distance
                NewListViewItem.SubItems.Add(DataReader.GetValue(2) & " (" & Form1.GetPercentage(DataReader.GetValue(2), CInt(DataReader.GetValue(2)) + CInt(DataReader.GetValue(3)) + CInt(DataReader.GetValue(4))) & ")") 'Demand Eco
                NewListViewItem.SubItems.Add(DataReader.GetValue(3) & " (" & Form1.GetPercentage(DataReader.GetValue(3), CInt(DataReader.GetValue(2)) + CInt(DataReader.GetValue(3)) + CInt(DataReader.GetValue(4))) & ")") 'Demand Bus
                NewListViewItem.SubItems.Add(DataReader.GetValue(4) & " (" & Form1.GetPercentage(DataReader.GetValue(4), CInt(DataReader.GetValue(2)) + CInt(DataReader.GetValue(3)) + CInt(DataReader.GetValue(4))) & ")") 'Demand First
                NewListViewItem.SubItems.Add(DataReader.GetValue(5) & " (" & Form1.GetPercentage(DataReader.GetValue(5), CInt(DataReader.GetValue(2))) & ")") 'Offer Eco
                NewListViewItem.SubItems.Add(DataReader.GetValue(6) & " (" & Form1.GetPercentage(DataReader.GetValue(6), CInt(DataReader.GetValue(3))) & ")") 'Offer Bus
                NewListViewItem.SubItems.Add(DataReader.GetValue(7) & " (" & Form1.GetPercentage(DataReader.GetValue(7), CInt(DataReader.GetValue(4))) & ")") 'Offer First
                NewListViewItem.SubItems.Add(CInt(DataReader.GetValue(2)) + CInt(DataReader.GetValue(3)) + CInt(DataReader.GetValue(4))) 'Overall Demand
                NewListViewItem.SubItems.Add(CInt(DataReader.GetValue(5)) + CInt(DataReader.GetValue(6)) + CInt(DataReader.GetValue(7)) & " (" & Form1.GetPercentage(CInt(DataReader.GetValue(5)) + CInt(DataReader.GetValue(6)) + CInt(DataReader.GetValue(7)), CInt(DataReader.GetValue(2)) + CInt(DataReader.GetValue(3)) + CInt(DataReader.GetValue(4))) & ")") 'Overall Offer
                Form1.LVRoutes.Items.Add(NewListViewItem)
            End While
        End If
        With Form1
            .New_TXTBX_RouteName.Text = ""
            .New_NUD_RouteDistance.Value = 0
            .New_NUD_DemandEco.Value = 0
            .New_NUD_DemandBus.Value = 0
            .New_NUD_DemandFirst.Value = 0
            .New_NUD_OfferEco.Value = 0
            .New_NUD_OfferBus.Value = 0
            .New_NUD_OfferFirst.Value = 0
        End With

    End Sub

    Public Sub WriteToDatabase(ArrayToWrite As String())
        Dim Query As String

        If ArrayToWrite(0) = "" Then
            MsgBox("Please enter a name for your route")
            Exit Sub
        End If
        Query = "SELECT * FROM Routes WHERE RouteName='" & ArrayToWrite(0) & "'"
        Dim DatabaseCommand As New OleDbCommand(Query, DatabaseConnection)
        DataReader = DatabaseCommand.ExecuteReader()
        If DataReader.HasRows Then
            MsgBox("Please choose a different name for your route!")
            Exit Sub
        End If

        Query = "INSERT INTO Routes (RouteName, RouteDistance, DemandEco, DemandBusiness, DemandFirst, OfferEco, OfferBusiness, OfferFirst) " &
            "VALUES ('" & ArrayToWrite(0) & "', " & ArrayToWrite(1) & ", " & ArrayToWrite(2) & ", " & ArrayToWrite(3) & ", " & ArrayToWrite(4) & ", " & ArrayToWrite(5) & ", " & ArrayToWrite(6) & ", " & ArrayToWrite(7) & ");"
        DatabaseCommand = New OleDbCommand(Query, DatabaseConnection)
        DatabaseCommand.ExecuteNonQuery()

        'Read Database again
        ReadDatabase("Routes", True)

    End Sub

    Public Sub CloseConnection()
        DatabaseConnection.Close()
    End Sub

    Public Sub DeleteFromDatabase(SearchString As String)
        Dim Query As String

        Query = "DELETE FROM Routes WHERE RouteName='" & SearchString & "'"
        Dim Command = New OleDbCommand(Query, DatabaseConnection)
        Command.ExecuteNonQuery()

        ReadDatabase("Routes")
    End Sub

End Class
