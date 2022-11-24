Class HomeTopBar
    Public Sub New()

        InitializeComponent()

        Dim iTopBar As vmHomeTopBar = New vmHomeTopBar(iState)
        Me.DataContext = iTopBar.ViewModelInstance
    End Sub

    Private Sub btnReports_Click(sender As Object, e As RoutedEventArgs)
        'iCurrentView.TopBarPage = New pgOrder_TopBar
        'iCurrentView.ContentPage = New pgOrder_Content
    End Sub
End Class
