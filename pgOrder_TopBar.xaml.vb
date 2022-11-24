Class pgOrder_TopBar
    Private Sub pgOrder_TopBar_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim iTopBar As New multiuse_FunctionBarViewModel(iState)
        Me.DataContext = iTopBar.ViewModelInstance
    End Sub
End Class
