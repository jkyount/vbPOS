﻿
Public Class ButtonClickCommand

    Inherits Commands.CommandBase

    Public Sub New(ViewModel As Object)
        MyBase.New(ViewModel)
    End Sub

    Public Overrides Sub Execute(parameter As Object)
        Me.ViewModel.ButtonClickCommandEvent()
    End Sub

End Class

