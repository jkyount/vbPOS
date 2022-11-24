



Imports System.Collections.ObjectModel
    Imports System.Windows


Public Class vmFloor
    Inherits ViewModelBase

    Public iFloorPlanViewModel As vmFloor
    Private pTablesInUse As Collection
    Private pServerTables As Collection
    Private pInUse As Visibility
    Private pOwnedByThisEmployee As Boolean
    Private pDisplayName As String
    Private pParentTable As String
    Private pFloorPlanCollection As ObservableCollection(Of vmFloor)
    Private pAccentColor As SolidColorBrush
    Private pTableClickCommand As TableClickCommand




    Public Function GetFloorPlanViewModel() As vmFloor
        Return iFloorPlanViewModel
    End Function

    Public Property FloorPlanCollection As ObservableCollection(Of vmFloor)
        Get
            Return pFloorPlanCollection
        End Get
        Set(value As ObservableCollection(Of vmFloor))
            pFloorPlanCollection = value

        End Set

    End Property
    Public Property TablesInUse As Collection
        Set(value As Collection)
            pTablesInUse = value
        End Set
        Get
            Return GetTablesInUse()
        End Get
    End Property

    Public Property ServerTables As Collection
        Set(value As Collection)
            pServerTables = value
        End Set
        Get
            Return pServerTables
        End Get
    End Property

    Property InUse() As Visibility
        Get
            Return pInUse
        End Get
        Set(value As Visibility)
            pInUse = value
            OnPropertyChanged(NameOf(InUse))
        End Set
    End Property

    Property OwnedByThisEmployee() As Boolean
        Get
            Return pOwnedByThisEmployee
        End Get
        Set(value As Boolean)
            pOwnedByThisEmployee = value
            OnPropertyChanged(NameOf(OwnedByThisEmployee))
        End Set
    End Property

    Property DisplayName() As String
        Get
            Return pDisplayName
        End Get
        Set(value As String)
            pDisplayName = value
            OnPropertyChanged(NameOf(DisplayName))
        End Set
    End Property

    Property ParentTable() As String
        Get
            Return pParentTable
        End Get
        Set(value As String)
            pParentTable = value
            OnPropertyChanged(NameOf(ParentTable))
        End Set
    End Property

    Property AccentColor() As SolidColorBrush
        Get
            Return pAccentColor
        End Get
        Set(value As SolidColorBrush)
            pAccentColor = value
            OnPropertyChanged(NameOf(AccentColor))
        End Set
    End Property

    Public Property TableClickCommand() As TableClickCommand
        Get
            TableClickCommand = pTableClickCommand
        End Get
        Set(value As TableClickCommand)
            pTableClickCommand = value
        End Set
    End Property




    Private Function GetVisibility(state As Boolean) As Visibility
        If state = True Then Return 0
        If state = False Then Return 1
        Return 1
    End Function

    Private Function GetAccentColor(ServerNum As Integer) As SolidColorBrush
        Dim brsh As New SolidColorBrush
        If ServerNum = iState.ThisEmployee.ServerNum Then
            brsh.Color = Colors.CornflowerBlue
            Return brsh
        End If
        brsh.Color = ColorConverter.ConvertFromString("#FFFFC357")
        Return brsh
    End Function



    Public Sub TableClickCommandEvent(table As Object)
        iState.ThisTable = New zclsTable(table.parent.parent.parent.name)
        iState.CurrentCheck = iState.ThisTable.Checks(1)
        iState.ThisOrder = New aclsOrder(iState.CurrentCheck)
        SyncValues()
        iCurrentView.ViewModel.SetContent(New pgOrder_TopBar, New pgOrder_Content)
    End Sub


    Public Sub New()
        Dim coll As Collection = GetTableStates()
        Dim collFloor As New ObservableCollection(Of vmFloor)
        For i As Integer = 1 To coll.Count
            collFloor.Add(New vmFloor(coll(i)))
        Next
        Me.FloorPlanCollection = collFloor
        MyBase.ViewModelInstance = Me
    End Sub

    Public Sub New(Val As Boolean, valll As Boolean)

    End Sub

    Private Sub New(TableDict As Dictionary(Of String, Object))
        Me.InUse = GetVisibility(TableDict("InUse"))
        Me.DisplayName = TableDict("DisplayName")
        Me.ParentTable = TableDict("ParentTable")
        Me.AccentColor = GetAccentColor(TableDict("ServerNum"))
        Me.TableClickCommand = New TableClickCommand(Me)

    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub



End Class
