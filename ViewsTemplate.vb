Public Class ViewsTemplate
    Private pView As Page
    Private pViewModel As ViewModelBase


    Public Property View As Page
        Set(value As Page)
            pView = value
        End Set
        Get
            Return pView
        End Get
    End Property

    Public Property ViewModel As ViewModelBase
        Set(value As ViewModelBase)
            pViewModel = value
        End Set
        Get
            Return pViewModel
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(View As Page, ViewModel As ViewModelBase)
        Me.View = View
        Me.ViewModel = View.DataContext
    End Sub
End Class
