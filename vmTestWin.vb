Public Class vmTestWin
    Inherits ViewModelBase

    Private pContent As ViewModelBase

    Public Property Content As ViewModelBase
        Set(value As ViewModelBase)
            pContent = value
        End Set
        Get
            Return pContent
        End Get
    End Property

    Public Sub New()
        Me.Content = New GridBaseViewModel
    End Sub

End Class
