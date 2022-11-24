Class HomeDynamicContent


    Public Sub New()

        InitializeComponent()

        Dim iContent As New vmHomeDynamicContent(New Floor)
        Me.DataContext = iContent.ViewModelInstance


    End Sub


End Class
