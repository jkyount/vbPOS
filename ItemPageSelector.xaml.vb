Public Class ItemPageSelector
    Public Sub New()
        InitializeComponent()
        Dim iItemPageSelectorViewModel As New ItemPageSelectorViewModel
        Me.DataContext = iItemPageSelectorViewModel
    End Sub
End Class
