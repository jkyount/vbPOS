Public Class OrderView
    Inherits ViewModelBase

    Private pActiveMenuMessenger As ActiveMenuMessenger
    Private pCheckView As UserControl
    Private pDynamicContent As UserControl


    Public Property ActiveMenuMessenger As ActiveMenuMessenger
        Set(value As ActiveMenuMessenger)
            pActiveMenuMessenger = value
        End Set
        Get
            Return pActiveMenuMessenger
        End Get
    End Property

    Public Property CheckView As UserControl
        Set(value As UserControl)
            pCheckView = value
        End Set
        Get
            Return pCheckView
        End Get
    End Property

    Public Property DynamicContent As UserControl
        Set(value As UserControl)
            pDynamicContent = value
        End Set
        Get
            Return pDynamicContent
        End Get
    End Property

    Public Sub New()
        Me.ActiveMenuMessenger = GetNewActiveMenuMessenger
        Me.ActiveMenuMessenger.InfoBar = New vmDefaultView
        'Me.CheckView = New CheckView
        Me.DynamicContent = New MenuItemSelection

    End Sub
    'Public Sub New()
    '    Me.ActiveMenuMessenger = iActiveMenuMessenger
    '    Me.ActiveMenuMessenger.InfoBar = New vmDefaultView
    '    Me.CheckView = New CheckView
    '    Me.DynamicContent = New MenuItemSelection

    'End Sub
    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
