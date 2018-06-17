Public Class Form1
    Dim DBHandler As New AccDBHandler

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Set Source
        DBHandler.SetSource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "\AirlinesManager.mdb")

        'Create New DB, if none found
        If Not (IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\AirlinesManager.mdb")) Then
            DBHandler.CreateNewDB()
        End If

        'Load Database
        DBHandler.ReadDatabase("Routes", True)
    End Sub

    Private Function GetPercentage(PercentageOf As Integer, SecondVal As Integer)
        Dim Percentage As Double

        Percentage = PercentageOf / SecondVal

        Return Percentage
    End Function

    Private Sub New_BTN_CreateNew_Click(sender As Object, e As EventArgs) Handles New_BTN_CreateNew.Click
        Dim StringToWrite(7) As String

        StringToWrite(0) = New_TXTBX_RouteName.Text
        StringToWrite(1) = New_NUD_RouteDistance.Value
        StringToWrite(2) = New_NUD_DemandEco.Value
        StringToWrite(3) = New_NUD_DemandBus.Value
        StringToWrite(4) = New_NUD_DemandFirst.Value
        StringToWrite(5) = New_NUD_OfferEco.Value
        StringToWrite(6) = New_NUD_OfferBus.Value
        StringToWrite(7) = New_NUD_OfferFirst.Value

        DBHandler.WriteToDatabase(StringToWrite)
    End Sub
End Class
