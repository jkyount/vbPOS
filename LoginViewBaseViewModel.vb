Public Class LoginViewBaseViewModel
    Inherits ViewModelBase

    Private pClockInCommand As ClockInCommand
    Private pLoginCommand As LoginCommand
    Private pExitCommand As ExitCommand
    Private pLoginID As Long

    Public Property ClockInCommand As ClockInCommand
        Set(value As ClockInCommand)
            pClockInCommand = value
        End Set
        Get
            Return pClockInCommand
        End Get
    End Property

    Public Property LoginCommand As LoginCommand
        Set(value As LoginCommand)
            pLoginCommand = value
        End Set
        Get
            Return pLoginCommand
        End Get
    End Property

    Public Property ExitCommand As ExitCommand
        Set(value As ExitCommand)
            pExitCommand = value
        End Set
        Get
            Return pExitCommand
        End Get
    End Property

    Public Property LoginID As Long
        Set(value As Long)
            pLoginID = value
        End Set
        Get
            Return pLoginID
        End Get
    End Property


    Public Sub New()
        Me.LoginID = 0
        Me.LoginCommand = New LoginCommand(Me)
        Me.ClockInCommand = New ClockInCommand(Me)
        Me.ExitCommand = New ExitCommand(Me)
    End Sub
    Public Sub ClockInCommandEvent()
        If IsNumeric(LoginID) = False Then Exit Sub
        If LoginID = 0 Then Exit Sub
        SetThisEmployee(LoginID)

        If IsLoginValid(LoginID) = False Then
            MsgBox("Invalid ID")
            LoginID = 0
            Exit Sub
        End If


        login_ClockIn_CLICK(LoginID)
        LoginID = 0
    End Sub

    Public Sub LoginCommandEvent()
        If IsNumeric(LoginID) = False Then
            MsgBox("Invalid ID")
            LoginID = 0
            Exit Sub
        End If
        If IsLoginValid(LoginID) = False Then
            MsgBox("Invalid ID")
            LoginID = 0
            Exit Sub
        End If
        Dim ThisEmployee As New zclsEmployee(LoginID)
        If ThisEmployee.ClockedIn = False Then
            MsgBox("Please clock in.")
            LoginID = 0
            Exit Sub
        End If
        login_Login_CLICK(LoginID)
    End Sub

    Public Sub ExitCommandEvent()
        CloseWindow("winLogin")
    End Sub


    Public Overrides Sub SetContent(Content1 As Page, Optional Content2 As Page = Nothing, Optional Content3 As Page = Nothing, Optional Content4 As Page = Nothing)
        Throw New NotImplementedException()
    End Sub
End Class
