Imports Scripting
Imports System.Collections.ObjectModel

Public Class MenuCategoriesViewModel
    Inherits ItemSelectionViewModelBase

    Private pButtonColl As ObservableCollection(Of MenuCategoryButton)
    'Private pButtonClickCommand As ButtonClickCommand

    Public Property ButtonColl As ObservableCollection(Of MenuCategoryButton)
        Set(value As ObservableCollection(Of MenuCategoryButton))
            pButtonColl = value
            OnPropertyChanged(NameOf(ButtonColl))
        End Set
        Get
            Return pButtonColl
        End Get
    End Property

    'Public Property ButtonClickCommand As ButtonClickCommand
    '    Set(value As ButtonClickCommand)
    '        pButtonClickCommand = value
    '    End Set
    '    Get
    '        Return pButtonClickCommand
    '    End Get
    'End Property

    Public Sub New()
        Me.ActiveMenuMessenger = iActiveMenuMessenger
        Me.ButtonColl = GetButtonColl()
        'Me.ButtonClickCommand = New ButtonClickCommand(Me)
    End Sub

    'Public Sub ButtonClickCommandEvent()
    '    ButtonColl(0).DataContext.DisplayName = "NEWNAME"
    'End Sub

    Private Function GetButtonColl() As ObservableCollection(Of MenuCategoryButton)
        Dim coll As Collection = GetFamilyGroup("MenuCategory")
        Dim BtnColl As New ObservableCollection(Of MenuCategoryButton)
        For Each member As Dictionary(Of String, Object) In coll
            BtnColl.Add(New MenuCategoryButton(New MenuCategoryButtonViewModel(member("ID"), member("DisplayName"))))
        Next
        Return BtnColl
    End Function

    Public Overrides Sub RefreshContent()

    End Sub
End Class
