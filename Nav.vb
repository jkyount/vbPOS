Module Nav

    Public Sub GoToOrder(LoginID As Long)
        EmployeeLogin(LoginID)
        Dim winOrder As New Order
        winOrder.OpenWin()
        CloseWindow("winLogin")
    End Sub

    Public Sub GoToHome()
        Dim winMain As New MainWindow
        winMain.InitializeComponent()

        winMain.OpenWin()
        CloseWindow("winLogin")
    End Sub

End Module
