
Imports System.Collections.ObjectModel

Public MustInherit Class vmBaseButton
    Inherits ViewModelBase
    Private pID As Integer
    Private pDisplayName As String
    Private pGUIInstance As Object
    Private pButtonClickCommand As ButtonClickCommand


    Public Property ButtonClickCommand As ButtonClickCommand
        Set(value As ButtonClickCommand)
            pButtonClickCommand = value
        End Set
        Get
            Return pButtonClickCommand
        End Get
    End Property

    Public Property ID As Integer
        Set(value As Integer)
            pID = value
        End Set
        Get
            Return pID
        End Get
    End Property

    Public Property DisplayName As String
        Set(value As String)
            pDisplayName = value
        End Set
        Get
            Return pDisplayName
        End Get
    End Property

    Public Property GUIInstance As Object
        Set(value As Object)
            pGUIInstance = value
        End Set
        Get
            Return pGUIInstance
        End Get
    End Property

    Public MustOverride Sub ButtonClickCommandEvent()

    Public MustOverride Function GetNew(ID As Integer, DisplayName As String)

    Public Sub New()

    End Sub

    Public Sub New(ID As Integer, DisplayName As String)
        Me.ID = ID
        Me.DisplayName = DisplayName
        Me.ButtonClickCommand = New ButtonClickCommand(Me)

    End Sub

    Public Function GetButtonColl(coll As Collection) As ObservableCollection(Of vmBaseButton)
        Dim BtnColl As New ObservableCollection(Of vmBaseButton)
        For Each key As Object In coll(1).keys
            BtnColl.Add(Me.GetNew(key, coll(1)(key)))
        Next
        Return BtnColl

    End Function

End Class


Public Class vmCondensedButton
    Inherits vmBaseButton

    Public Overrides Sub ButtonClickCommandEvent()
        iActiveMenuMessenger.InfoBar = New ParentSelectorViewModel(Me.DisplayName)
        Ssfub(Me.ID)

    End Sub

    Public Sub New()
        MyBase.New
    End Sub

    Public Sub New(ID As Integer, DisplayName As String)
        MyBase.New(ID, DisplayName)

    End Sub

    Public Overrides Function GetNew(ID As Integer, DisplayName As String) As Object
        Return New vmCondensedButton(ID, DisplayName)
    End Function
End Class

Public Class vmExpandedButton
    Inherits vmBaseButton

    Public Overrides Sub ButtonClickCommandEvent()

        iActiveMenuMessenger.InfoBar = New ParentSelectorViewModel(Me.DisplayName)
        Ssfub(Me.ID)

    End Sub

    Public Sub New()
        MyBase.New
    End Sub
    Public Sub New(ID As Integer, DisplayName As String)
        MyBase.New(ID, DisplayName)
    End Sub

    Public Overrides Function GetNew(ID As Integer, DisplayName As String) As Object
        Return New vmExpandedButton(ID, DisplayName)
    End Function
End Class


Public Class vmComponentButton
    Inherits vmBaseButton

    Public Overrides Sub ButtonClickCommandEvent()
        iActiveMenuMessenger.InfoBar = New ParentSelectorViewModel(Me.DisplayName)
        Ssfub(Me.ID)

    End Sub

    Public Sub New()
        MyBase.New
    End Sub
    Public Sub New(ID As Integer, DisplayName As String)
        MyBase.New(ID, DisplayName)

    End Sub

    Public Overrides Function GetNew(ID As Integer, DisplayName As String) As Object
        Return New vmComponentButton(ID, DisplayName)
    End Function
End Class





'Public MustInherit Class ButtonViewModelBase
'    Inherits ViewModelBase
'    Private pID As Integer
'    Private pDisplayName As String
'    Private pGUIInstance As Object
'    Private pButtonClickCommand As ButtonClickCommand

'    Public Property ButtonClickCommand As ButtonClickCommand
'        Set(value As ButtonClickCommand)
'            pButtonClickCommand = value
'            OnPropertyChanged(NameOf(ButtonClickCommand))
'        End Set
'        Get
'            Return pButtonClickCommand
'        End Get
'    End Property

'    Public Property ID As Integer
'        Set(value As Integer)
'            pID = value
'            OnPropertyChanged(NameOf(ID))
'        End Set
'        Get
'            Return pID
'        End Get
'    End Property

'    Public Property DisplayName As String
'        Set(value As String)
'            pDisplayName = value
'            OnPropertyChanged(NameOf(DisplayName))
'        End Set
'        Get
'            Return pDisplayName
'        End Get
'    End Property

'    Public Property GUIInstance As Object
'        Set(value As Object)
'            pGUIInstance = value
'        End Set
'        Get
'            Return pGUIInstance
'        End Get
'    End Property

'    Public MustOverride Sub ButtonClickCommandEvent()

'    Public Sub New(ID As Integer, DisplayName As String)
'        Me.ID = ID
'        Me.DisplayName = DisplayName
'        Me.ButtonClickCommand = New ButtonClickCommand(Me)
'    End Sub

'    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
'        Throw New NotImplementedException()
'    End Sub
'End Class


'Public Class CondensedMenuButtonViewModel
'    Inherits ButtonViewModelBase

'    Public Overrides Sub ButtonClickCommandEvent()
'        MsgBox("Implement Condensed Menu Button Click Logic")
'    End Sub

'    Public Sub New(ID As Integer, DisplayName As String)
'        MyBase.New(ID, DisplayName)
'    End Sub

'End Class

'Public Class ExpandedMenuButtonViewModel
'    Inherits ButtonViewModelBase

'    Public Overrides Sub ButtonClickCommandEvent()
'        MsgBox("Implement Extended Menu Button Click Logic")
'    End Sub

'    Public Sub New(ID As Integer, DisplayName As String)
'        MyBase.New(ID, DisplayName)
'    End Sub
'End Class

