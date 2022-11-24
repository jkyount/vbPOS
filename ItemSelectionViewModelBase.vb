Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public MustInherit Class ItemSelectionViewModelBase
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged


    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub


    Private pActiveMenuMessenger As ActiveMenuMessenger
    Private pInfoBarConstructor As InfoBarConstructor
    Private pTopRow As ItemSelectionViewModelBase
    Private pSubRow As ItemSelectionViewModelBase

    'Private pCanvasRow As IButtonCanvas

    'Public Property CanvasRow As IButtonCanvas
    '    Set(value As IButtonCanvas)
    '        pCanvasRow = value
    '        OnPropertyChanged(NameOf(CanvasRow))
    '    End Set
    '    Get
    '        Return pCanvasRow
    '    End Get
    'End Property


    Public Property TopRow As ItemSelectionViewModelBase
        Set(value As ItemSelectionViewModelBase)
            pTopRow = value
            OnPropertyChanged(NameOf(TopRow))
        End Set
        Get
            Return pTopRow
        End Get
    End Property

    Public Property SubRow As ItemSelectionViewModelBase
        Set(value As ItemSelectionViewModelBase)
            pSubRow = value
            OnPropertyChanged(NameOf(SubRow))
        End Set
        Get
            Return pSubRow
        End Get
    End Property

    Public Property ActiveMenuMessenger As ActiveMenuMessenger
        Set(value As ActiveMenuMessenger)
            pActiveMenuMessenger = value
            OnPropertyChanged(NameOf(ActiveMenuMessenger))
        End Set
        Get
            Return pActiveMenuMessenger
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

    Public Sub New()
        Me.ActiveMenuMessenger = iActiveMenuMessenger

    End Sub


    Overridable Sub RefreshContent()

    End Sub


End Class

Public Class ItemSelectionArchetype
    Inherits ItemSelectionViewModelBase




    Public Sub New()
        MyBase.New
        Me.ActiveMenuMessenger.InfoBarConstructor.TopRow = New MenuCategoriesViewModel
        Me.ActiveMenuMessenger.InfoBarConstructor.SubRow = New ItemPageSelectorViewModel
        'Me.CanvasRow = New BlankMenuButtonCanvasViewModel
    End Sub



End Class
