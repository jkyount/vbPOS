Public Class multiuse_FunctionBarViewModel
    Inherits ViewModelBase

    Private pServerName As String
    Private pOrderName As String
    Private pLogoutCommand As LogoutCommand
    'Private pGoToPaymentScreenCommand As GoToPaymentScreenCommand
    'Private pLogoutCommand As LogoutCommand

    'Private pCancelOrderCommand As CancelOrderCommand

    Public Property ServerName As String
        Set(value As String)
            pServerName = value
        End Set
        Get
            Return pServerName
        End Get
    End Property

    Public Property OrderName As String
        Set(value As String)
            pOrderName = value
        End Set
        Get
            Return pOrderName
        End Get
    End Property



    'Public Property SendOrderCommand() As SendOrderCommand
    '    Get
    '        SendOrderCommand = pSendOrderCommand
    '    End Get
    '    Set(value As SendOrderCommand)
    '        pSendOrderCommand = value
    '    End Set
    'End Property

    'Public Property GoToPaymentScreenCommand() As GoToPaymentScreenCommand
    '    Get
    '        GoToPaymentScreenCommand = pGoToPaymentScreenCommand
    '    End Get
    '    Set(value As GoToPaymentScreenCommand)
    '        pGoToPaymentScreenCommand = value
    '    End Set
    'End Property

    Public Property LogoutCommand() As LogoutCommand
        Get
            LogoutCommand = pLogoutCommand
        End Get
        Set(value As LogoutCommand)
            pLogoutCommand = value
        End Set
    End Property

    'Public Property CancelOrderCommand() As CancelOrderCommand
    '    Get
    '        CancelOrderCommand = pCancelOrderCommand
    '    End Get
    '    Set(value As CancelOrderCommand)
    '        pCancelOrderCommand = value
    '    End Set
    'End Property

    Public Sub LogoutCommandEvent()
        order_Logout_CLICK()

    End Sub
    Public Sub OpenTablesCommandEvent()
        home_OpenTables_CLICK()
    End Sub

    Public Sub ViewFloorCommandEvent()
        home_ViewFloor_CLICK()
    End Sub
    Public Sub New(iState As State)
        Me.LogoutCommand = New LogoutCommand(Me)
        'Me.OpenTablesCommand = New OpenTablesCommand(Me)
        'Me.ViewFloorCommand = New ViewFloorCommand(Me)

        Me.ServerName = iState.ThisEmployee.FirstName
        Me.OrderName = iState.ThisOrder.ValueDict("OrderName")
        MyBase.ViewModelInstance = Me
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
