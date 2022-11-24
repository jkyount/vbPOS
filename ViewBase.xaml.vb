'Class ViewBase

'    Private pViewModel As ViewModelBase
'    Private pView As Page

'    Public Property ViewModel As ViewModelBase
'        Set(value As ViewModelBase)
'            pViewModel = value
'        End Set
'        Get
'            Return pViewModel
'        End Get
'    End Property

'    Public Property View As Page
'        Set(value As Page)
'            pView = value
'        End Set
'        Get
'            Return pView
'        End Get
'    End Property
'    'Public Sub New(ContentPage As Page, TopBarPage As Page)

'    '    ' This call is required by the designer.
'    '    InitializeComponent()

'    '    ' Add any initialization after the InitializeComponent() call.
'    '    Dim iViewBaseViewModel As New ViewBaseViewModel(ContentPage, TopBarPage)

'    '    Me.DataContext = iViewBaseViewModel
'    '    'iCurrentView.CurrentViewBaseViewModel = iViewBaseViewModel.ViewModelInstance
'    'End Sub

'    'Private Sub ViewBase_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
'    '    Me.DataContext = iCurrentView
'    'End Sub
'    Public Sub New()

'        ' This call is required by the designer.
'        InitializeComponent()

'        ' Add any initialization after the InitializeComponent() call.

'    End Sub
'    Public Sub New(GivenView As ViewBase)

'        ' This call is required by the designer.
'        InitializeComponent()

'        ' Add any initialization after the InitializeComponent() call.
'        Me.ViewModel = GivenView.ViewModel
'        Me.View = GivenView.View
'        Me.DataContext = Me.ViewModel
'        iCurrentView.View = Me.View

'    End Sub

'    'Private Sub ViewBase_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
'    '    Me.DataContext = iCurrentView
'    'End Sub
'End Class
