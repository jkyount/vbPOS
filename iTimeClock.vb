Module iTimeClock


    Public Sub TimeClock_ClockIn(ID As Long)
        Dim iTimeClock As New aclsTimeClock
        iTimeClock.ClockIn(ID)
        iTimeClock = Nothing
    End Sub

    Public Sub TimeClock_ClockOut(ID As Long)
        Dim iTimeClock As New aclsTimeClock
        iTimeClock.ClockOut(ID)
        iTimeClock = Nothing
    End Sub
End Module
