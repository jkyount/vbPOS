Public Class MenuItemSelection
    Public Sub New()
        InitializeComponent()

        Dim iMenuItemSelectionViewModel As New ItemSelectionArchetype
        Me.DataContext = iMenuItemSelectionViewModel

    End Sub
End Class
