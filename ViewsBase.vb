Public Class ViewsBase

    Private pView As Page
    Private pViewModel As ViewModelBase

    Public Property ViewModel As ViewModelBase
        Set(value As ViewModelBase)
            pViewModel = value
        End Set
        Get
            Return pViewModel
        End Get
    End Property


    Public Property View As Page
        Set(value As Page)
            pView = value
        End Set
        Get
            Return pView
        End Get
    End Property


    Public Sub New(GivenView As ViewsTemplate)
        Me.View = GivenView.View
        Me.ViewModel = GivenView.ViewModel
    End Sub


End Class
