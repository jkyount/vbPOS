Public Class vmHomeDynamicContent
    Inherits ViewModelBase

    Public iHomeContent As vmHomeDynamicContent
    Private pContentPage As Page





    Private pSwitchUserCommand As SwitchUserCommand
    Private pOpenTablesCommand As OpenTablesCommand
    Private pViewFloorCommand As ViewFloorCommand

    Public Property SwitchUserCommand() As SwitchUserCommand
        Get
            SwitchUserCommand = pSwitchUserCommand
        End Get
        Set(value As SwitchUserCommand)
            pSwitchUserCommand = value
        End Set
    End Property

    Public Property OpenTablesCommand() As OpenTablesCommand
        Get
            OpenTablesCommand = pOpenTablesCommand
        End Get
        Set(value As OpenTablesCommand)
            pOpenTablesCommand = value
        End Set
    End Property

    Public Property ViewFloorCommand() As ViewFloorCommand
        Get
            ViewFloorCommand = pViewFloorCommand
        End Get
        Set(value As ViewFloorCommand)
            pViewFloorCommand = value
        End Set
    End Property

    Public Sub SwitchUserCommandEvent()
        GoToLoginScreen()
    End Sub


    Public Sub OpenTablesCommandEvent()


        DirectCast(MyBase.ViewModelInstance, vmHomeDynamicContent).ContentPage = New OpenChecksList
        ''iHomeContent.ContentPage = New pgHome_OpenTables
    End Sub

    Public Sub ViewFloorCommandEvent()
        DirectCast(MyBase.ViewModelInstance, vmHomeDynamicContent).ContentPage = New Floor

        'iHomeContent.ContentPage = New pgHome_FloorPlan
    End Sub
    Public Sub New(ContentPage As Page)
        Me.SwitchUserCommand = New SwitchUserCommand(Me)
        Me.OpenTablesCommand = New OpenTablesCommand(Me)
        Me.ViewFloorCommand = New ViewFloorCommand(Me)
        Me.ContentPage = ContentPage
        'iHomeContent = Me
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

    Public Function GetHomeContentViewModel()

        Return iHomeContent
    End Function

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
