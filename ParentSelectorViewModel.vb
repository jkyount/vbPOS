Public Class ParentSelectorViewModel
    Inherits ItemSelectionViewModelBase

    Public Sub New(ParentName As String)
        MyBase.New
        Me.ActiveMenuMessenger.CurrentParentName = "Selecting Options For: " & ParentName
    End Sub
End Class
