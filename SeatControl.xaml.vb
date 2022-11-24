Public Class SeatControl
    Public Sub New()

        InitializeComponent()
        Dim iSeatControlViewModel As New SeatControlViewModel(iState.CurrentSeat)
        Me.DataContext = iSeatControlViewModel


    End Sub
End Class
