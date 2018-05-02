Public Class Form1
    Dim DBHandler As New AccDBHandler
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DBHandler.SetSource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Files.mdb")
        DBHandler.CreateNewDB(True)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DBHandler.SetSource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Files.mdb")
        DBHandler.ReadDatabase("Files", True)
        DataGridView1.DataSource = DBHandler.DataSetCollection.Tables(0) 'Debug
        DataGridView1.Refresh() 'Debug
    End Sub
End Class
