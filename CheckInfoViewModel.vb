Public Class CheckInfoViewModel
    Inherits ViewModelBase

    Private pValueName As String
    Private pValue As Decimal
    Private pThisOrder As aclsOrder

    Public ReadOnly Property SubTotal As String
        Get
            Return FormatCurrency(ThisOrder.ValueDict("SubTotal"), 2)
        End Get
    End Property

    Public ReadOnly Property Tax As String
        Get
            Return FormatCurrency(ThisOrder.ValueDict("Tax"), 2)
        End Get
    End Property

    Public ReadOnly Property Total As String
        Get
            Return FormatCurrency(ThisOrder.ValueDict("Total"), 2)
        End Get
    End Property

    Public Property ThisOrder As aclsOrder
        Set(value As aclsOrder)
            pThisOrder = value
            OnPropertyChanged(NameOf(ThisOrder))
        End Set
        Get
            Return pThisOrder
        End Get
    End Property

    Public Property ValueName As String
        Set(value As String)
            pValueName = value
        End Set
        Get
            Return pValueName
        End Get
    End Property

    Public Property Value As Decimal
        Set(value As Decimal)
            pValue = value
        End Set
        Get
            Return pValue
        End Get
    End Property

    Public Sub New(iState As State)
        Me.ThisOrder = iState.ThisOrder
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
