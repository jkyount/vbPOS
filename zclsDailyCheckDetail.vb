Public Class zclsDailyCheckDetail
    Inherits DBObject

    Public dbDailyCheckDetail As New ADODB.Connection
    Public rsDailyCheckDetail As New ADODB.Recordset









    Public Overrides Function GetDb() As String
        GetDb = "DailyCheckDetail"
    End Function

    Public Overrides Function GetDbFile() As String
        GetDbFile = "CheckDb"
    End Function

    Public Function GetArchive() As String
        GetArchive = "ArchivedCheckDetail"
    End Function
    Public Function GetArchiveDbFile() As String
        GetArchiveDbFile = "ReportsDb"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbDailyCheckDetail
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsDailyCheckDetail
    End Function

    Public Function GetNextLocalGroup(check As String) As Integer
        OpenDbs()
        rsDailyCheckDetail.let_Source("SELECT LocalGroup FROM DailyCheckDetail WHERE CheckNumber = """ & check & """ ORDER BY LocalGroup ASC")
        rsDailyCheckDetail.Open()

        If rsDailyCheckDetail.RecordCount = 0 Then
            GetNextLocalGroup = 1
            rsDailyCheckDetail.Close()
            dbDailyCheckDetail.Close()
            Exit Function
        End If
        rsDailyCheckDetail.MoveLast()
        GetNextLocalGroup = rsDailyCheckDetail.Fields("LocalGroup").Value + 1
        CloseDbs()
    End Function

    Public Function GetNextEntityGroup(check As String) As Integer
        OpenDbs()
        rsDailyCheckDetail.let_Source("SELECT EntityGroup FROM DailyCheckDetail WHERE CheckNumber = """ & check & """ ORDER BY EntityGroup ASC")
        rsDailyCheckDetail.Open()

        If rsDailyCheckDetail.RecordCount = 0 Then
            GetNextEntityGroup = 1
            rsDailyCheckDetail.Close()
            dbDailyCheckDetail.Close()
            Exit Function
        End If
        rsDailyCheckDetail.MoveLast()
        GetNextEntityGroup = rsDailyCheckDetail.Fields("EntityGroup").Value + 1
        CloseDbs()
    End Function

    Public Sub SendCurrentItems(check As String)
        pSendCurrentItems(check)
    End Sub

    Private Sub pSendCurrentItems(check As String)
        OpenDbs()
        rsDailyCheckDetail.let_Source("SELECT * From DailyCheckDetail WHERE CheckNumber = """ & iState.CurrentCheck & """ ORDER BY Seat ASC")
        rsDailyCheckDetail.Open()
        Do Until rsDailyCheckDetail.EOF
            rsDailyCheckDetail.Fields("Sent").Value = True
            rsDailyCheckDetail.MoveNext()
        Loop
        CloseDbs()
    End Sub

    Public Sub AddCurrentItems(coll As Collection)
        Dim member As aclsItem
        Dim LocalGroup As Integer, EntityGroup As Integer
        EntityGroup = GetNextEntityGroup(iState.CurrentCheck)
        For Each member In coll
            LocalGroup = GetNextLocalGroup(iState.CurrentCheck)
            AddItem(member, LocalGroup, EntityGroup)
        Next member
    End Sub


    Public Sub AddItem(item As aclsItem, LocalGroup As Integer, EntityGroup As Integer)
        OpenDbs()
        rsDailyCheckDetail.let_Source("SELECT * From DailyCheckDetail WHERE CheckNumber = """ & iState.CurrentCheck & """ ORDER BY Seat ASC")
        rsDailyCheckDetail.Open()
        rsDailyCheckDetail.AddNew()
        rsDailyCheckDetail.Fields("CheckNumber").Value = iState.CurrentCheck
        rsDailyCheckDetail.Fields("ItemID").Value = item.ItemID
        rsDailyCheckDetail.Fields("ItemIndicator").Value = item.ItemIndicator
        rsDailyCheckDetail.Fields("ItemName").Value = item.ItemName
        rsDailyCheckDetail.Fields("Price").Value = item.Price
        rsDailyCheckDetail.Fields("Family").Value = item.Family
        rsDailyCheckDetail.Fields("Category").Value = item.Category
        rsDailyCheckDetail.Fields("AlwaysTax").Value = item.AlwaysTax
        rsDailyCheckDetail.Fields("PrintKitchen").Value = item.PrintKitchen
        rsDailyCheckDetail.Fields("PrintPantry").Value = item.PrintPantry
        rsDailyCheckDetail.Fields("Seat").Value = GetCurrentSeat()
        rsDailyCheckDetail.Fields("LocalGroup").Value = LocalGroup
        rsDailyCheckDetail.Fields("IsPrimaryItem").Value = item.IsPrimaryItem
        rsDailyCheckDetail.Fields("EntityGroup").Value = EntityGroup
        rsDailyCheckDetail.Update()
        CloseDbs()
    End Sub
End Class


