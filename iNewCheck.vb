Module iNewCheck
    Public Function InitializeCheck(check As String) As String
        Dim iNewCheck As New zclsNewCheck
        iNewCheck.InitializeCheck(check)
        InitializeCheck = check
    End Function

    Public Function GetNextCheck() As String
        Dim iNewCheck As New zclsNewCheck
        Return iNewCheck.GetNextCheck
    End Function

    Public Function PeekNextCheck() As String
        Dim iNewCheck As New zclsNewCheck
        Return iNewCheck.PeekNextCheck
    End Function


    Public Sub SetCheckInUse(check As String)
        Dim iNewCheck As New zclsNewCheck
        iNewCheck.SetCheckInUse(check)
    End Sub

    Public Sub SetCheckUnused(check As String)
        Dim iNewCheck As New zclsNewCheck
        iNewCheck.SetCheckUnused(check)
    End Sub

End Module
