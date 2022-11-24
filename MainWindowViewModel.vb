Public Class MainWindowViewModel
    Inherits ViewModelBase
    Private piCurrentView

    Public Property View() As CurrentView
        Get
            Return piCurrentView
        End Get
        Set(value As CurrentView)
            piCurrentView = value
            OnPropertyChanged(NameOf(View))
        End Set
    End Property

    Public Sub New()
        Me.View = iCurrentView
        MyBase.ViewModelInstance = Me
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
