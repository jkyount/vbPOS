Class pgOrder_Content

    Public Sub New()
        InitializeComponent()
        Dim iOrderView As New OrderView
        Me.DataContext = iOrderView

    End Sub
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

    End Sub
End Class
