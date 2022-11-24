Imports Scripting

Module iEmployee

    Public Sub SetThisEmployee(ID As Long)
        iState.ThisEmployee = New zclsEmployee(ID)

        'ThisEmployee.Reset()
        'ThisEmployee.IDNumber = ID
    End Sub
    Public Function GetEmployeeObj(ID As Long)
        Dim iEmployee As zclsEmployee
        iEmployee = New zclsEmployee(ID)

        GetEmployeeObj = iEmployee
        iEmployee = Nothing
    End Function
    Public Function IsLoginValid(IDNumber As Long) As Boolean
        Dim iEmployee As New zclsEmployee(IDNumber)
        IsLoginValid = iEmployee.IsLoginValid(IDNumber)
        iEmployee = Nothing
    End Function

    Public Sub AddEmployeeJob(JobCode As Integer, PayRate As Double, ServerNum As Integer)
        Dim iEmployee As New zclsEmployee(0)
        iEmployee.AddNewJob(JobCode, PayRate, ServerNum)
        iEmployee = Nothing
    End Sub

    Public Sub RemoveEmployeeJob(JobCode As Integer, ServerNum As Integer)
        Dim iEmployee As New zclsEmployee(0)
        iEmployee.RemoveJob(JobCode, ServerNum)
        iEmployee = Nothing
    End Sub

    Public Function GetJobDict(ServerNum As Integer) As Dictionary(Of String, Object)
        Dim iEmployee As New zclsEmployee(0)
        GetJobDict = iEmployee.GetJobDict(ServerNum)
        iEmployee = Nothing
    End Function
End Module
