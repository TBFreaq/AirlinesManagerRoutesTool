Imports System.ComponentModel

Public Class Form1
    Dim DBHandler As New AccDBHandler
    Dim OldName As String

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Set Source
        DBHandler.SetSource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "\AirlinesManager.mdb")

        'Create New DB, if none found
        If Not (IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\AirlinesManager.mdb")) Then
            Dim TableName As String = "Routes"
            Dim ColumnNames() As String = {"RouteName", "RouteDistance", "DemandEco", "DemandBus", "DemandFirst", "OfferEco", "OfferBus", "OfferFirst"}
            DBHandler.CreateNewDB(TableName, ColumnNames)
        End If

        'Load Database
    End Sub

    Public Function GetPercentage(PercentageOf As Integer, SecondVal As Integer, Optional ReturnNA As Boolean = True)
        Dim Percentage As Double
        If (SecondVal = 0) Then
            If ReturnNA = True Then
                Return "N/A"
            Else
                Return -1
            End If
        Else
            Percentage = PercentageOf / SecondVal
            Return Format(Percentage * 100, "00.00")
        End If
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

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        DBHandler.CloseConnection()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        DBHandler.ReadDatabase("Routes", True)
    End Sub

    Private Sub Edit_BTN_DeleteRoute_Click(sender As Object, e As EventArgs) Handles Edit_BTN_DeleteRoute.Click

        'Check if a route is selected, then delete it
        If (LVRoutes.SelectedItems.Count = 0) Then
            MsgBox("Please select a Route")
        Else
            DBHandler.DeleteFromDatabase(LVRoutes.SelectedItems(0).Text)
        End If
    End Sub

    Private Sub LVRoutes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LVRoutes.SelectedIndexChanged
        Try
            Edit_TXTBX_RouteName.Text = LVRoutes.SelectedItems(0).Text
            OldName = LVRoutes.SelectedItems(0).Text
            Edit_NUD_RouteDistance.Value = CInt(LVRoutes.SelectedItems(0).SubItems(1).Text) 'Distance
            Edit_NUD_DemandEco.Value = Microsoft.VisualBasic.Left(LVRoutes.SelectedItems(0).SubItems(2).Text, InStr(LVRoutes.SelectedItems(0).SubItems(2).Text, "(") - 1) 'Demand Eco
            Edit_NUD_DemandBus.Value = Microsoft.VisualBasic.Left(LVRoutes.SelectedItems(0).SubItems(3).Text, InStr(LVRoutes.SelectedItems(0).SubItems(3).Text, "(") - 1) 'Demand Bus
            Edit_NUD_DemandFirst.Value = Microsoft.VisualBasic.Left(LVRoutes.SelectedItems(0).SubItems(4).Text, InStr(LVRoutes.SelectedItems(0).SubItems(4).Text, "(") - 1) 'Demand First
            Edit_NUD_OfferEco.Value = Microsoft.VisualBasic.Left(LVRoutes.SelectedItems(0).SubItems(5).Text, InStr(LVRoutes.SelectedItems(0).SubItems(5).Text, "(") - 1) 'Offer Eco
            Edit_NUD_OfferBus.Value = Microsoft.VisualBasic.Left(LVRoutes.SelectedItems(0).SubItems(6).Text, InStr(LVRoutes.SelectedItems(0).SubItems(6).Text, "(") - 1) 'Offer Business
            Edit_NUD_OfferFirst.Value = Microsoft.VisualBasic.Left(LVRoutes.SelectedItems(0).SubItems(7).Text, InStr(LVRoutes.SelectedItems(0).SubItems(7).Text, "(") - 1) 'Offer First
        Catch ex As Exception
            'For some reason it works
        End Try
        'Calculate Seats
        CalcSeats()
    End Sub

    Private Sub CalcSeats()
        If (LVRoutes.SelectedIndices.Count = 0) Then Exit Sub

        Dim DemandEco As Double
        Dim DemandBus As Double
        Dim DemandFirst As Double
        Dim SeatsEco As Integer
        Dim SeatsBus As Integer
        Dim SeatsFirst As Integer
        Dim TotalDemand As Integer
        Dim TotalSeats As Integer
        Dim Divisor As Integer

        'Seat Calculation
        '1 First = 4 Eco // 1 Bus = 2 Eco
        TotalDemand = Edit_NUD_DemandEco.Value + Edit_NUD_DemandBus.Value + Edit_NUD_DemandFirst.Value
        TotalSeats = Seats_NUD_MaxSeats.Value
        DemandEco = Edit_NUD_DemandEco.Value
        DemandBus = Edit_NUD_DemandBus.Value
        DemandFirst = Edit_NUD_DemandFirst.Value
        Divisor = (DemandEco * 1) + (DemandBus * 2) + (DemandFirst * 4)
        SeatsEco = TotalSeats * DemandEco / Divisor
        SeatsBus = TotalSeats * DemandBus / Divisor
        SeatsFirst = TotalSeats * DemandFirst / Divisor

        Seats_TXTBX_SeatsEco.Text = SeatsEco
        Seats_TXTBX_SeatsBus.Text = SeatsBus
        Seats_TXTBX_SeatsFirst.Text = SeatsFirst

    End Sub

    Private Sub Seats_NUD_MaxSeats_ValueChanged(sender As Object, e As EventArgs) Handles Seats_NUD_MaxSeats.ValueChanged
        CalcSeats()
    End Sub

    Private Sub BTNEditEntry_Click(sender As Object, e As EventArgs) Handles BTNEditEntry.Click

        'Check if a row is selected
        If (LVRoutes.SelectedItems.Count = 0) Then
            MsgBox("Please select a route!")
            Exit Sub
        End If

        'Edit the selected route by deleting it and making a new one
        DBHandler.DeleteFromDatabase(OldName)
        Dim StringToWrite(7) As String

        StringToWrite(0) = Edit_TXTBX_RouteName.Text
        StringToWrite(1) = Edit_NUD_RouteDistance.Value
        StringToWrite(2) = Edit_NUD_DemandEco.Value
        StringToWrite(3) = Edit_NUD_DemandBus.Value
        StringToWrite(4) = Edit_NUD_DemandFirst.Value
        StringToWrite(5) = Edit_NUD_OfferEco.Value
        StringToWrite(6) = Edit_NUD_OfferBus.Value
        StringToWrite(7) = Edit_NUD_OfferFirst.Value

        'Write the new Row
        DBHandler.WriteToDatabase(StringToWrite)

    End Sub

    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
    End Sub
End Class
