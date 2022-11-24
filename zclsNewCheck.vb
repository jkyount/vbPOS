Public Class zclsNewCheck
    Inherits DBObject
    Public dbNewCheck As New ADODB.Connection
    Public rsNewCheck As New ADODB.Recordset

    Public Overrides Function GetDb() As String
        GetDb = "DailyCheckIndex"
    End Function
    Public Overrides Function GetDbFile() As String
        GetDbFile = "CheckDb"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbNewCheck
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsNewCheck
    End Function



    Public Function GetNextCheck() As String
        GetNextCheck = pGetNextCheck()
    End Function

    Public Function pGetNextCheck() As String
        OpenDbs()
        rsNewCheck.Source = "CheckNumbers"
        rsNewCheck.Open()

        Do Until rsNewCheck.Fields("Used").Value = False
            rsNewCheck.MoveNext()

            If rsNewCheck.EOF Then

                rsNewCheck.Close()
                dbNewCheck.Close()
                pGetNextCheck = ResetChecks()
                Exit Function
            End If
        Loop
        pGetNextCheck = "ch" & rsNewCheck.Fields("CheckNumber").Value
        rsNewCheck.Fields("Used").Value = True
        rsNewCheck.Update()
        rsNewCheck.Close()
        dbNewCheck.Close()
    End Function
    Private Function ResetChecks() As String
        OpenDbs()
        rsNewCheck.Source = "CheckNumbers"
        rsNewCheck.Open()
        rsNewCheck.MoveFirst()

        Do Until rsNewCheck.EOF
            rsNewCheck.Fields("Used").Value = False
            rsNewCheck.MoveNext()
        Loop
        rsNewCheck.UpdateBatch()
        rsNewCheck.Close()
        dbNewCheck.Close()
        ResetChecks = pGetNextCheck()


    End Function

    Public Function PeekNextCheck() As String
        PeekNextCheck = pPeekNextCheck()
    End Function

    Public Function pPeekNextCheck() As String
        OpenDbs()
        rsNewCheck.Source = "CheckNumbers"
        rsNewCheck.Open()

        Do Until rsNewCheck.Fields("Used").Value = False
            rsNewCheck.MoveNext()
        Loop
        pPeekNextCheck = "ch" & rsNewCheck.Fields("CheckNumber").Value
        rsNewCheck.Close()
        dbNewCheck.Close()
    End Function

    Public Sub SetCheckInUse(check As String)
        pSetCheckInUse(check)
    End Sub
    Private Sub pSetCheckInUse(check As String)

        OpenDbs()
        rsNewCheck.Source = "SELECT * FROM CheckNumbers WHERE CheckNumber = """ & check & """"
        rsNewCheck.Open()
        rsNewCheck.Fields("Used").Value = True
        rsNewCheck.Update()
        rsNewCheck.Close()
        dbNewCheck.Close()
    End Sub

    Public Sub SetCheckUnused(check As String)

        pSetCheckUnused(check)
    End Sub

    Private Sub pSetCheckUnused(check As String)
        Dim CheckInt As Integer
        CheckInt = CInt(Right(check, (Len(check) - 2)))
        OpenDbs()
        rsNewCheck.Source = "SELECT * FROM CheckNumbers WHERE CheckNumber = " & CheckInt & ""
        rsNewCheck.Open()
        rsNewCheck.Fields("Used").Value = False
        rsNewCheck.Update()
        rsNewCheck.Close()
        dbNewCheck.Close()
    End Sub

    Public Sub InitializeCheck(check As String)
        pInitializeCheck(check)
    End Sub
    Private Sub pInitializeCheck(check As String)
        OpenDbs()
        rsNewCheck.Source = "DailyCheckIndex"
        rsNewCheck.Open()
        rsNewCheck.AddNew()
        rsNewCheck.Fields("CheckNumber").Value = check
        rsNewCheck.Update()
        rsNewCheck.Close()
        dbNewCheck.Close()
    End Sub
End Class
