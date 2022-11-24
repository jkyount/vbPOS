Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Window
Imports POS1
Module Functions_Login

    Public Sub EmployeeLogin(LoginID As Long)
        iState.ThisEmployee = New zclsEmployee(LoginID)


    End Sub

    Public Sub EmployeeClockIn(LoginID As Long)
        If Not ThisEmployee.ClockedIn = False Then
            Select Case MsgBox("Clock out?", vbYesNo)
                Case vbYes
                    ClockOut(LoginID)
                    Exit Sub
                Case vbNo
                    Exit Sub
            End Select
        End If
        ClockIn(LoginID)
    End Sub
    Public Sub ClockIn(ID As Long)
        ThisEmployee.ClockIn(ID)
        TimeClock_ClockIn(ID)
    End Sub

    Public Sub ClockOut(ID As Long)
        ThisEmployee.ClockOut(ID)
        TimeClock_ClockOut(ID)
    End Sub
End Module
