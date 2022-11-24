










Public Class SwitchUserCommand

    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.SwitchUserCommandEvent()
    End Sub

End Class

Public Class TableClickCommand

    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.TableClickCommandEvent(parameter)
    End Sub

End Class

Public Class OpenTablesCommand

    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.OpenTablesCommandEvent()
    End Sub

End Class

Public Class ViewFloorCommand

    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.ViewFloorCommandEvent()
    End Sub

End Class





