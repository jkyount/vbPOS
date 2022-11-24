Module CLICK
    Public Sub login_Login_CLICK(LoginID As Long)
        EmployeeLogin(LoginID)
        GoToHome()
    End Sub

    Public Sub login_ClockIn_CLICK(LoginID As Long)
        EmployeeClockIn(LoginID)
    End Sub


    Public Sub GoToLoginScreen()
        Dim winLogin As New Login
        winLogin.OpenWin()
        CloseWindow("winMain")

    End Sub

    Public Sub home_OpenTables_CLICK()



    End Sub

    Public Sub home_ViewFloor_CLICK()
        Dim OpenTablesView As StackPanel = GetWindows("winMain").FindName("OpenTablesStackPanel")
        OpenTablesView.Visibility = Visibility.Hidden
    End Sub

    Public Sub order_Logout_CLICK()
        Dim winLogin As New Login
        winLogin.OpenWin()
        CloseWindow("winMain")
    End Sub
End Module
