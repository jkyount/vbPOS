
Public Class CurrentView

    Inherits ViewModelBase
    Private pViewModel As ViewModelBase

    Public Property ViewModel As ViewModelBase
        Set(value As ViewModelBase)
            pViewModel = value
            OnPropertyChanged(NameOf(ViewModel))
        End Set
        Get
            Return pViewModel
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(InitialContent As ViewModelBase)
        Me.ViewModel = InitialContent
        MyBase.ViewModelInstance = Me
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class




'Public Class CurrentView


'    Inherits ViewModelBase


'    Private pView As Page
'    Private pInitialPage As Page
'    Private pViewModel As ViewModelBase


'    Public Property ViewModel As ViewModelBase
'        Set(value As ViewModelBase)
'            pViewModel = value

'            OnPropertyChanged(NameOf(ViewModel))
'        End Set
'        Get
'            Return pViewModel
'        End Get
'    End Property


'    Public Property View As Page
'        Set(value As Page)
'            pView = value

'            OnPropertyChanged(NameOf(View))
'        End Set
'        Get
'            Return pView
'        End Get
'    End Property

'    Public Sub New()

'    End Sub

'    Public Sub New(InitialContent As ViewsBase)
'        Me.View = InitialContent.View
'        Me.ViewModel = InitialContent.ViewModel
'        MyBase.ViewModelInstance = Me
'    End Sub

'    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
'        Throw New NotImplementedException()
'    End Sub
'End Class
