Imports System.Collections.ObjectModel
Imports System.Windows
Public Class vmServerFunctions
    Inherits ViewModelBase

    Private pSwitchUserCommand As SwitchUserCommand
    Private pOpenTablesCommand As OpenTablesCommand
    Private pViewFloorCommand As ViewFloorCommand

    Public Property SwitchUserCommand() As SwitchUserCommand
        Get
            SwitchUserCommand = pSwitchUserCommand
        End Get
        Set(value As SwitchUserCommand)
            pSwitchUserCommand = value
        End Set
    End Property

    Public Property OpenTablesCommand() As OpenTablesCommand
        Get
            OpenTablesCommand = pOpenTablesCommand
        End Get
        Set(value As OpenTablesCommand)
            pOpenTablesCommand = value
        End Set
    End Property

    Public Property ViewFloorCommand() As ViewFloorCommand
        Get
            ViewFloorCommand = pViewFloorCommand
        End Get
        Set(value As ViewFloorCommand)
            pViewFloorCommand = value
        End Set
    End Property

    Public Sub SwitchUserCommandEvent()
        GoToLoginScreen()
    End Sub


    Public Sub OpenTablesCommandEvent()
        home_OpenTables_CLICK()
    End Sub

    Public Sub ViewFloorCommandEvent()
        home_ViewFloor_CLICK()
    End Sub
    Public Sub New()
        Me.SwitchUserCommand = New SwitchUserCommand(Me)
        Me.OpenTablesCommand = New OpenTablesCommand(Me)
        Me.ViewFloorCommand = New ViewFloorCommand(Me)
    End Sub

End Class

