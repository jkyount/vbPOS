Public Class SeatControlViewModel
    Inherits State

    Private pCurrentSeat As Integer
    Private pSeatIndicatorText As String
    Private pAdvanceSeatCommand As AdvanceSeatCommand
    Private pResetSeatCommand As ResetSeatCommand




    Public Overloads Property CurrentSeat As Integer
        Get
            Return pCurrentSeat
        End Get
        Set(value As Integer)
            pCurrentSeat = value

        End Set
    End Property

    Public Property SeatIndicatorText As String
        Set(value As String)
            pSeatIndicatorText = value
            OnPropertyChanged(NameOf(SeatIndicatorText))
        End Set
        Get
            Return pSeatIndicatorText
        End Get
    End Property

    Public Property AdvanceSeatCommand As AdvanceSeatCommand
        Set(value As AdvanceSeatCommand)
            pAdvanceSeatCommand = value
        End Set
        Get
            Return pAdvanceSeatCommand
        End Get
    End Property

    Public Property ResetSeatCommand As ResetSeatCommand
        Set(value As ResetSeatCommand)
            pResetSeatCommand = value
        End Set
        Get
            Return pResetSeatCommand
        End Get
    End Property


    Public Sub New(CurrentSeat As Integer)
        Me.CurrentSeat = CurrentSeat
        Me.SeatIndicatorText = "Entry for Seat " & Me.CurrentSeat
        Me.AdvanceSeatCommand = New AdvanceSeatCommand(Me)
        Me.ResetSeatCommand = New ResetSeatCommand(Me)
    End Sub

    Public Sub AdvanceSeatCommandEvent()
        SyncSeat(CurrentSeat + 1)
    End Sub

    Public Sub ResetSeatCommandEvent()
        SyncSeat(1)
    End Sub

    Public Sub SyncSeat(Seat As Integer)
        Me.CurrentSeat = Seat
        MyBase.CurrentSeat = Seat
        Me.SeatIndicatorText = "Entry for Seat " & Seat
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
