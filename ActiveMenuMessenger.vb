Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class ActiveMenuMessenger
    Inherits ViewModelBase
    Private pActiveMenuCategory As Integer
    Private pActiveMenuName As String
    Private pRenderTransform As TranslateTransform
    Private pMenuConstructor As MenuConstructor
    Private pGrid As UserControl
    Private pCurrentParentName As String
    Private pInfoBar As ItemSelectionViewModelBase
    Private pCheckView As ViewModelBase

    Public Property CheckView As ViewModelBase
        Set(value As ViewModelBase)
            pCheckView = value
            OnPropertyChanged(NameOf(CheckView))
        End Set
        Get
            Return pCheckView
        End Get
    End Property

    Public Property InfoBar As ItemSelectionViewModelBase
        Set(value As ItemSelectionViewModelBase)
            pInfoBar = value
            OnPropertyChanged(NameOf(InfoBar))
        End Set
        Get
            Return pInfoBar
        End Get
    End Property

    Public Property CurrentParentName As String
        Set(value As String)
            pCurrentParentName = value
            OnPropertyChanged(NameOf(CurrentParentName))
        End Set
        Get
            Return pCurrentParentName
        End Get
    End Property

    Private pInfoBarConstructor As InfoBarConstructor

    Public Property MenuConstructor As MenuConstructor
        Set(value As MenuConstructor)
            pMenuConstructor = value
            OnPropertyChanged(NameOf(MenuConstructor))
        End Set
        Get
            Return pMenuConstructor
        End Get
    End Property

    Public Property InfoBarConstructor As InfoBarConstructor
        Set(value As InfoBarConstructor)
            pInfoBarConstructor = value
            OnPropertyChanged(NameOf(InfoBarConstructor))
        End Set
        Get
            Return pInfoBarConstructor
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

    Public Property RenderTransform As TranslateTransform
        Set(value As TranslateTransform)
            pRenderTransform = value
            OnPropertyChanged(NameOf(RenderTransform))
        End Set
        Get
            Return pRenderTransform
        End Get
    End Property

    Public Property ActiveMenuCategory As Integer
        Set(value As Integer)
            pActiveMenuCategory = value
            Refresh(pActiveMenuCategory)
            OnPropertyChanged(NameOf(ActiveMenuCategory))
        End Set
        Get
            Return pActiveMenuCategory
        End Get
    End Property

    Public Property ActiveMenuName As String
        Set(value As String)
            pActiveMenuName = value
            OnPropertyChanged(NameOf(ActiveMenuName))
        End Set
        Get
            Return pActiveMenuName
        End Get
    End Property

    Public Sub New()
        Me.MenuConstructor = New MenuConstructor
        Me.InfoBarConstructor = New InfoBarConstructor
        Me.CheckView = New vmCheckView
    End Sub

    Private Sub Refresh(FamilyID As Integer)
        Dim iFamily As New zclsFamily(FamilyID)
        Me.MenuConstructor.Refresh(iFamily)
        Me.CheckView = New vmCheckView
    End Sub

    Public Sub Reset()
        Me.MenuConstructor = New MenuConstructor
        Me.InfoBar = New vmDefaultView
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class



