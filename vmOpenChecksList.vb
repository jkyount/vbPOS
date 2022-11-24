Imports System.Collections.ObjectModel

Public Class vmOpenChecksList
    Inherits ViewModelBase

    'Public iOpenTablesViewModel As OpenTablesViewModel

    Private pCheck As String
    Private pOrderName As String
    Private pTotal As String
    Private pOwner As String
    Private pOpenTablesCollection As ObservableCollection(Of vmOpenChecksList)

    'Public Function GetOpenTablesViewModel() As OpenTablesViewModel
    '    Return iOpenTablesViewModel
    'End Function

    Public Sub New(Check As String, OrderName As String, Total As String, Owner As String)
        Me.Check = Check
        Me.OrderName = OrderName
        Me.Total = "$" & Total
        Me.Owner = Owner
    End Sub

    Public Sub New(Checks As Array)
        Me.OpenTablesCollection = GetListContentsCollection(Checks)
        MyBase.ViewModelInstance = Me
    End Sub

    Public Property OpenTablesCollection As ObservableCollection(Of vmOpenChecksList)
        Get
            Return pOpenTablesCollection
        End Get
        Set(value As ObservableCollection(Of vmOpenChecksList))
            pOpenTablesCollection = value
        End Set

    End Property

    Property Check() As String
        Get
            Return pCheck
        End Get
        Set(value As String)
            pCheck = value
            OnPropertyChanged(NameOf(Check))
        End Set
    End Property

    Property OrderName() As String
        Get
            Return pOrderName
        End Get
        Set(value As String)
            pOrderName = value
            OnPropertyChanged(NameOf(OrderName))
        End Set
    End Property

    Property Total() As String
        Get
            Return pTotal
        End Get
        Set(value As String)
            pTotal = value
            OnPropertyChanged(NameOf(Total))
        End Set
    End Property

    Property Owner() As String
        Get
            Return pOwner
        End Get
        Set(value As String)
            pOwner = value
            OnPropertyChanged(NameOf(Owner))
        End Set
    End Property

    Public Function GetListContentsCollection(Checks As Array) As ObservableCollection(Of vmOpenChecksList)

        Dim OpenTable As vmOpenChecksList
        Dim coll As ObservableCollection(Of vmOpenChecksList)
        coll = New ObservableCollection(Of vmOpenChecksList)
        For i As Integer = 0 To UBound(Checks)
            OpenTable = New vmOpenChecksList(Checks(i, 0), Checks(i, 1), FormatNumber(Checks(i, 3), 2), Checks(i, 4))
            coll.Add(OpenTable)
        Next
        Return coll
    End Function

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
