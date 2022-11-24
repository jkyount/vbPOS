
Imports Scripting

Public Class zclsEmployee
    Inherits DBObject
    Private pFirstName As String
    Private pLastName As String
    Private pIDNumber As Long
    Private pClockedIn As Boolean
    Private pServerNum As Integer
    Private pAccentColor As Brush

    Public dbEmployee As New ADODB.Connection
    Public rsEmployee As New ADODB.Recordset

    Public Enum JobVals
        EmptyJob = 7
    End Enum


    Public Overrides Function GetDb() As String
        GetDb = "Employee"
    End Function

    Public Overrides Function GetDbFile() As String
        GetDbFile = "Employee"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbEmployee
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsEmployee
    End Function



    Public Property AccentColor() As Brush
        Set(value As Brush)
            pAccentColor = value
        End Set
        Get
            Dim brsh As New SolidColorBrush
            brsh.Color = ColorConverter.ConvertFromString(ToString(Me.GetEmployeeParam("AccentColor")))
            'Return Color Me.GetEmployeeParam("AccentColor")
            Return brsh
        End Get
    End Property


    Public Property FirstName() As String
        Set(value As String)
            pFirstName = value
        End Set
        Get
            Return Me.GetEmployeeParam("FirstName")
        End Get
    End Property


    Public Property ClockedIn() As Boolean
        Set(value As Boolean)
            pClockedIn = value
        End Set
        Get
            Return Me.GetEmployeeParam("ClockedIn")
        End Get
    End Property

    Public Property LastName() As String
        Set(value As String)
            pLastName = value
        End Set
        Get
            Return Me.GetEmployeeParam("LastName")
        End Get
    End Property

    Public Property IDNumber() As Long
        Set(value As Long)
            pIDNumber = value
        End Set
        Get
            Return pIDNumber

        End Get
    End Property
    Public Property ServerNum() As Long
        Set(value As Long)
            pServerNum = value
        End Set
        Get
            Return Me.GetEmployeeParam("ServerNum")
        End Get
    End Property


    '==========================================================================

    Public Sub New(ID As Long)
        Me.IDNumber = ID
    End Sub

    Public Function GetAccentColor() As String
        OpenDbs()
        rsEmployee.let_Source("SELECT AccentColor FROM Employee WHERE ServerNum = " & Me.ServerNum & "")
        rsEmployee.Open()

        If IsDBNull(rsEmployee.Fields("AccentColor").Value) Then
            GetAccentColor = "255, 195, 87"
            rsEmployee.Close()
            dbEmployee.Close()
            Exit Function
        End If
        GetAccentColor = ToString(rsEmployee.Fields("AccentColor").Value)
        rsEmployee.Close()
        dbEmployee.Close()
    End Function

    Public Sub SetAccentColor(color As String)
        Update(Wrap(GetNewUpdateObj("ServerNum", Me.ServerNum, "AccentColor", color)))
    End Sub


    Public Function GetEmployeeParam(param As String) As Object
        OpenDbs()
        rsEmployee.let_Source("SELECT * FROM Employee WHERE IDNumber = " & Me.IDNumber & "")
        rsEmployee.Open()
        GetEmployeeParam = rsEmployee.Fields(param).Value
        rsEmployee.Close()
        dbEmployee.Close()
    End Function

    Public Function IsLoginValid(IDNumber As Long) As Boolean
        IsLoginValid = pIsLoginValid(IDNumber)
    End Function
    Private Function pIsLoginValid(IDNumber As Long) As Boolean
        OpenDbs()
        rsEmployee.let_Source("Employee")
        rsEmployee.Open()

        Do
            If rsEmployee.Fields("IDNumber").Value = IDNumber Then
                pIsLoginValid = True
                rsEmployee.Close()
                dbEmployee.Close()
                Exit Function
            End If
            rsEmployee.MoveNext()
        Loop Until rsEmployee.EOF
        pIsLoginValid = False
        rsEmployee.Close()
        dbEmployee.Close()
    End Function

    Public Function Reset()
        Me.FirstName = ""
        Me.IDNumber = 0
        Me.LastName = ""
        Me.ServerNum = 0
    End Function

    Public Sub ClockIn(ID As Long)
        Dim iEmployee As New zclsEmployee(0)

        Update(iEmployee.Wrap(GetNewUpdateObj("IDNumber", ID, "ClockedIn", True)))
    End Sub

    Public Sub ClockOut(ID As Long)
        Dim iEmployee As New zclsEmployee(0)
        Update(iEmployee.Wrap(GetNewUpdateObj("IDNumber", ID, "ClockedIn", False)))
    End Sub

    Public Sub ClockOutAll()
        Dim iEmployee As New zclsEmployee(0)
        Update(iEmployee.Wrap(GetNewUpdateObj("ClockedIn", True, "ClockedIn", False)))
    End Sub

    Public Function GetEmployees() As Object
        Dim arr As Object

        arr = FilteredMatch(Wrap(GetNewMatchObj("NOT ServerNum", 0)), {"ServerNum"}, {"FirstName"}, {"LastName"})

        Dim i As Integer
        For i = 1 To UBound(arr)
            arr(i)(0, 0) = arr(i)(0, 0) & " - " & arr(i)(1, 0) & " " & arr(i)(2, 0)
        Next i
        GetEmployees = arr

    End Function

    Public Function GetAllJobs() As Object
        Dim iDataObj As New aclsDataObject
        iDataObj = Wrap(GetNewMatchObj("NOT ID", 0))
        iDataObj.Db = "Job"

        GetAllJobs = FilteredMatch(iDataObj, {"Job"})
    End Function

    Public Function GetJobs(ServerNum As Integer) As Object

        GetJobs = FilteredMatch(Wrap(GetNewMatchObj("ServerNum", ServerNum)), {"Job1"}, {"Job2"}, {"Job3"})

    End Function

    Public Function JobCodeToName(JobCode As Object) As String
        Dim iDataObj As New aclsDataObject
        iDataObj = Wrap(GetNewMatchObj("ID", JobCode))
        iDataObj.Db = "Job"
        JobCodeToName = ValueMatch(iDataObj, "Job")
        iDataObj = Nothing
    End Function

    'Public Function FormatGetJobs(Jobs As Object) As Object
    '    Dim arr As Object
    '    ReDim arr(1 To UBound(Jobs(1)) + 1)
    '    Dim arr2(0, 0) As Object
    '    Dim i As Integer
    '    For i = 1 To UBound(Jobs(1)) + 1
    '        arr2(0, 0) = Jobs(1)(i - 1, 0)
    '        arr(i) = arr2
    '    Next i
    '    FormatGetJobs = arr
    'End Function

    Public Function JobNameToCode(Job As String) As Integer
        Dim iDataObj As aclsDataObject = Wrap(Wrap(GetNewMatchObj("Job", Job)))
        iDataObj.Db = "Job"
        JobNameToCode = ValueMatch(iDataObj, "ID")
        iDataObj = Nothing
    End Function

    Public Sub AddNewJob(JobCode As Integer, PayRate As Double, ServerNum As Integer)
        Dim NewJobIndex As String
        NewJobIndex = GetJobIndex(ServerNum, JobVals.EmptyJob)
        Update(Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Job" & NewJobIndex, JobCode)))
        Update(Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Payrate" & NewJobIndex, PayRate)))
    End Sub

    Public Sub RemoveJob(JobCode As Integer, ServerNum As Integer)
        Dim JobIndex As String
        JobIndex = GetJobIndex(ServerNum, JobCode)
        Update(Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Job" & JobIndex, JobVals.EmptyJob)))
        Update(Wrap(GetNewUpdateObj("ServerNum", ServerNum, "Payrate" & JobIndex, 0)))
    End Sub

    Public Function GetJobDict(ServerNum As Integer) As Dictionary(Of String, Object)
        Dim dict As New Dictionary(Of String, Object)
        dict = GetValueDict(Wrap(GetNewMatchObj("ServerNum", ServerNum)))(1)
        Dim JobDict As New Dictionary(Of String, Object)
        Dim JunkDict As Dictionary(Of String, Object)
        Dim i As Integer
        For i = 1 To 3
            JunkDict = New Dictionary(Of String, Object)
            JunkDict.Add("Job", dict("Job" & i))
            JunkDict.Add("Payrate", dict("Payrate" & i))
            JobDict.Add("Job" & i, JunkDict)
        Next i
        GetJobDict = JobDict
        dict = Nothing
        JobDict = Nothing
        JunkDict = Nothing
    End Function

    Private Function GetJobIndex(ServerNum As Integer, TargetJobCode As Integer) As String
        Dim JobDict As New Dictionary(Of String, Object)
        JobDict = GetJobDict(ServerNum)

        For Each key As Object In JobDict
            If JobDict(key)("Job") = TargetJobCode Then
                GetJobIndex = Right(key, 1)
                Exit For
                Exit Function
            End If
        Next key
    End Function
End Class
