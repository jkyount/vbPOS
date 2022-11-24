Class OpenChecksList



    Private Sub pgHome_OpenTables_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim iContent As New vmOpenChecksList(GetChecks("Closed", False))
        Me.DataContext = iContent.ViewModelInstance
    End Sub
End Class
