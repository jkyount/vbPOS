

Public Class ClockInCommand
    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.ClockInCommandEvent()
    End Sub

End Class

Public Class LoginCommand
    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.LoginCommandEvent()
    End Sub

End Class

Public Class ExitCommand
    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.ExitCommandEvent()
    End Sub

End Class

