Public Class State
    Inherits ViewModelBase

    Private pThisEmployee As zclsEmployee
    Private pThisTable As zclsTable
    Private pCurrentCheck As String
    Private pThisOrder As aclsOrder
    Private pCurrentSeat As Integer

    Public Property ThisEmployee As zclsEmployee
        Set(value As zclsEmployee)
            pThisEmployee = value
        End Set
        Get
            Return pThisEmployee
        End Get
    End Property

    Public Property ThisTable As zclsTable
        Set(value As zclsTable)
            pThisTable = value
        End Set
        Get
            Return pThisTable
        End Get
    End Property

    Public Property ThisOrder As aclsOrder
        Set(value As aclsOrder)
            pThisOrder = value
        End Set
        Get
            Return pThisOrder
        End Get
    End Property

    Public Property CurrentCheck As String
        Set(value As String)
            pCurrentCheck = value
        End Set
        Get
            Return pCurrentCheck
        End Get
    End Property

    Public Property CurrentSeat As Integer
        Set(value As Integer)
            pCurrentSeat = value
            OnPropertyChanged(NameOf(CurrentSeat))
        End Set
        Get
            Return pCurrentSeat
        End Get
    End Property




    Public Sub New()
        'Me.ThisEmployee = GetThisEmployee()
        'Me.ThisTable = GetThisTable()
        'Me.CurrentCheck = GetCurrentCheck()

        Me.CurrentSeat = 1
        'Me.ThisOrder = GetThisOrder(Me)
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
