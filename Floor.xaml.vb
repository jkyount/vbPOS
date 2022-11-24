Class Floor


    Public Sub New()
        InitializeComponent()

    End Sub

    Private Sub Floor_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim Tbl As ctrlTableGroup
        Dim iContent As New vmFloor()
        Dim iFloorPlanViewModel As vmFloor = iContent.GetFloorPlanViewModel
        iFloorPlanViewModel = iContent




        For Each member As Object In iFloorPlanViewModel.FloorPlanCollection
            Tbl = Me.FindName(member.ParentTable)
            Tbl.DataContext = member
        Next member
    End Sub
End Class
