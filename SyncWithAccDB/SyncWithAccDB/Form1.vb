Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DBHanlder As New AccDBHandler
        DBHanlder.SetSource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Files.mdb")
        DBHanlder.CreateNewDB(True)
    End Sub
End Class
