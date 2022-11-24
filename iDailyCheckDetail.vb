Module iDailyCheckDetail
    Public Sub SendCurrentItems(check As String)

        Dim iDailyCheckDetail As New zclsDailyCheckDetail
        iDailyCheckDetail.SendCurrentItems(check)
        iDailyCheckDetail = Nothing

    End Sub

    Public Function GetNextLocalGroup(check As String) As Integer
        Dim iDailyCheckDetail As New zclsDailyCheckDetail
        GetNextLocalGroup = iDailyCheckDetail.GetNextLocalGroup(check)
        iDailyCheckDetail = Nothing
    End Function

    Public Function GetNextEntityGroup(check As String) As Integer
        Dim iDailyCheckDetail As New zclsDailyCheckDetail
        GetNextEntityGroup = iDailyCheckDetail.GetNextEntityGroup(check)
        iDailyCheckDetail = Nothing
    End Function
End Module
