Public Class CheckInfo
    Public Sub New()


        InitializeComponent()
        Dim iCheckInfoViewModel As New CheckInfoViewModel(iState)
        Me.DataContext = iCheckInfoViewModel


    End Sub
End Class
