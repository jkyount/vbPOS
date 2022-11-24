Imports System.Collections.ObjectModel
Public Class MenuConstructor

    Inherits ViewModelBase
    Private pMenuStyle As aclsMenuStyle
    Private pGrid As UserControl
    Private pButtonType As vmBaseButton
    Private pButtonColl As ObservableCollection(Of vmBaseButton)

    Public Property MenuStyle As aclsMenuStyle
        Set(value As aclsMenuStyle)
            pMenuStyle = value
        End Set
        Get
            Return pMenuStyle
        End Get
    End Property

    Public Property Grid As UserControl
        Set(value As UserControl)
            pGrid = value
            OnPropertyChanged(NameOf(Grid))
        End Set
        Get
            Return pGrid
        End Get
    End Property

    Public Property ButtonType As vmBaseButton
        Set(value As vmBaseButton)
            pButtonType = value
        End Set
        Get
            Return pButtonType
        End Get
    End Property

    Public Property ButtonColl As ObservableCollection(Of vmBaseButton)
        Set(value As ObservableCollection(Of vmBaseButton))
            pButtonColl = value

        End Set
        Get
            Return pButtonColl
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ID As Integer)

    End Sub

    'Public Sub Refresh(ID As Integer)
    '    'Dim iFamily As New zclsFamily(ID)
    '    Me.Grid = Me.MenuStyle.MenuComponents.Grid
    '    Me.ButtonType = Me.MenuStyle.MenuComponents.ButtonType
    '    Me.ButtonColl = GetButtonColl(ID)

    '    Me.Grid.DataContext = Me
    'End Sub

    Public Sub Refresh(iFamily As zclsFamily)
        'Dim iFamily As New zclsFamily(ID)
        Me.MenuStyle = iFamily.MenuStyleObj
        Me.Grid = Me.MenuStyle.MenuComponents.Grid
        Me.ButtonType = Me.MenuStyle.MenuComponents.ButtonType
        Me.ButtonColl = GetButtonColl(iFamily.FamilyID)

        Me.Grid.DataContext = Me
    End Sub

    Public Sub New(iFamily As zclsFamily, iMenuStyle As aclsMenuStyle)

    End Sub

    'Public Function GetGrid(iFamily As zclsFamily) As UserControl
    '    If iFamily.IsMultiMenu = True Then
    '        MsgBox("Implement MultiMenu GUI")
    '        Return New CondensedGrid(Me)
    '        Exit Function
    '    End If
    '    If iFamily.Count = 0 Then Return Nothing
    '    If iFamily.MenuStyle = MenuStyleEnum.Condensed Then
    '        Return New CondensedGrid(Me)
    '    End If
    '    If iFamily.MenuStyle = MenuStyleEnum.Expanded Then
    '        Return New ExpandedGrid(Me)
    '    End If
    '    Return Nothing
    'End Function

    'Public Function GetButtonType(iFamily As zclsFamily) As vmBaseButton
    '    If iFamily.MenuStyle = MenuStyleEnum.Condensed Then
    '        Return New vmCondensedButton
    '    End If
    '    If iFamily.MenuStyle = MenuStyleEnum.Expanded Then
    '        Return New vmExpandedButton
    '    End If
    '    Return Nothing
    'End Function

    Private Function GetButtonColl(FamilyID As Integer) As ObservableCollection(Of vmBaseButton)
        Dim coll As Collection = GetFamilyMembers(FamilyID)
        Return Me.ButtonType.GetButtonColl(coll)
    End Function





    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
