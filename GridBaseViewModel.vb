Public Class GridBaseViewModel
    Inherits ViewModelBase

    Private pMenuConstructor As MenuConstructor


    Public Property MenuConstructor As MenuConstructor
        Set(value As MenuConstructor)
            pMenuConstructor = value
            OnPropertyChanged(NameOf(MenuConstructor))

        End Set
        Get
            Return pMenuConstructor
        End Get
    End Property

    Public Sub New()
        Me.MenuConstructor = New MenuConstructor

    End Sub
    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
