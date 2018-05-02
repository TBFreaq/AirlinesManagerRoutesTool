Public Class AccDBHandler

    Dim Source As String = ""
    Dim FileName As String = ""
    Dim CompletePath As String = ""

    Public Sub SetSource(SourceStr As String, FileNameStr As String)
        Source = SourceStr
        FileName = FileNameStr
        CompletePath = Source & FileName
    End Sub

    Public Sub CreateNewDB(Optional Overwrite As Boolean = False)

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

        ADOXTable.Name = "Files"
        ADOXTable.Columns.Append("FileSource", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable.Columns.Append("FileSourceDateTimeCreated", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable.Columns.Append("FileSourceMD5", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable.Columns.Append("FileDestination", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable.Columns.Append("FileDestinationDateTimeCreated", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable.Columns.Append("FileDestinationMD5", ADOX.DataTypeEnum.adVarWChar)

        ADOXCatalog.Tables.Append(ADOXTable)
        ADOXTable.Indexes.Append(ADOXIndex)

        ADOXTable2.Name = "Statistics"
        ADOXTable2.Columns.Append("LastAccess", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable2.Columns.Append("FilesChanged", ADOX.DataTypeEnum.adVarWChar)
        ADOXTable2.Columns.Append("NumberOfFiles", ADOX.DataTypeEnum.adVarWChar)

        ADOXCatalog.Tables.Append(ADOXTable2)
        ADOXTable2.Indexes.Append(ADOXIndex)

        ADOXTable = Nothing
        ADOXTable2 = Nothing
        ADOXCatalog = Nothing
        ADOXIndex = Nothing

    End Sub

    Public Sub ReadDatabase()



    End Sub

    Public Sub WriteToDatabase(ArrayToWrite As String)



    End Sub

End Class
