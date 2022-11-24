Public Class ViewBaseViewModel
    Inherits ViewModelBase

    Public iViewBaseViewModel As ViewBaseViewModel
    Private pContentPage As Page
    Private pTopBarPage As Page


    Public Sub New()

    End Sub

    Public Sub New(ContentPage As Page, TopBarPage As Page)

        Me.ContentPage = ContentPage
        Me.TopBarPage = TopBarPage
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

    'Public Sub SetContent(ContentPage As Page, TopBarPage As Page)
    '    Me.ContentPage = ContentPage
    '    Me.TopBarPage = TopBarPage
    'End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Me.ContentPage = Content2
        Me.TopBarPage = Content1
    End Sub


End Class
