Public Class MenuCategoryButton
    Private pID As Integer

    Public Property ID As Integer
        Set(value As Integer)
            pID = value
        End Set
        Get
            Return pID
        End Get
    End Property

    Public Sub New(ViewModel As MenuCategoryButtonViewModel)
        'Implement ButtonViewModelBase with properties of ID and DisplayName
        InitializeComponent()
        Me.ID = ViewModel.ID
        Me.Name = "Btn" & Me.ID
        Me.DataContext = ViewModel
        ViewModel.GUIInstance = Me

    End Sub
End Class
