Public Class vmHomeTopBar

    Inherits ViewModelBase


    Private pServerName As String


    Public Property ServerName As String
        Set(value As String)
            pServerName = value
            OnPropertyChanged(NameOf(ServerName))
        End Set
        Get
            Return pServerName

        End Get
    End Property


    Public Sub New(iState As State)


        Me.ServerName = iState.ThisEmployee.FirstName
        MyBase.ViewModelInstance = Me
    End Sub

    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
