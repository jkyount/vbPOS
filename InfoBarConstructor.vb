Public Class InfoBarConstructor
    Inherits ViewModelBase

    Private pTopRow As ItemSelectionViewModelBase
    Private pSubRow As ItemSelectionViewModelBase

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

    Public Sub New()

    End Sub


End Class
