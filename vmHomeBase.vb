Public Class vmHomeBase
    Inherits ViewModelBase


    Private pContentPage As Page
    Private pTopBarPage As Page


    Public Sub New()
        Me.ContentPage = New HomeDynamicContent
        Me.TopBarPage = New HomeTopBar
        MyBase.ViewModelInstance = Me
    End Sub


    Public Property ContentPage As Page
        Set(value As Page)
            pContentPage = value
            OnPropertyChanged(NameOf(ContentPage))

        End Set
        Get
            Return pContentPage
        End Get
    End Property

    Public Property TopBarPage As Page
        Set(value As Page)
            pTopBarPage = value
            OnPropertyChanged(NameOf(TopBarPage))

        End Set
        Get
            Return pTopBarPage
        End Get
    End Property



    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Me.TopBarPage = Content1
        Me.ContentPage = Content2
    End Sub
End Class
