Module iTable
    Public ThisTable As New zclsTable

    Public Sub SetTableInUse(check As String, state As String)
        Dim iTable As New zclsTable
        iTable.SetTableInUse(check, state)
        iTable = Nothing
    End Sub

    Public Function GetTablesInUse() As Collection
        Dim iTable As New zclsTable
        GetTablesInUse = iTable.GetTablesInUse
        iTable = Nothing
    End Function

    Public Function GetServerTables(ServerNum As Integer) As Collection
        Dim iTable As New zclsTable
        GetServerTables = iTable.GetServerTables(ServerNum)
        iTable = Nothing
    End Function

    Public Function GetTableStates() As Collection
        Dim iTable As New zclsTable
        GetTableStates = iTable.GetTableStates()
        iTable = Nothing
    End Function

    Public Function GetNextTable(ParentTable As String) As String
        Dim iTable As New zclsTable
        GetNextTable = iTable.GetNextTable(ParentTable)
        iTable = Nothing
    End Function

    Public Sub ClearTableStates()
        Dim iTable As New zclsTable
        iTable.ClearTableStates()
        iTable = Nothing
    End Sub

    Public Sub SetTableState(ParentTable As String, Table As String)
        ThisTable.ParentTable = ParentTable
        ThisTable.Table = Table
    End Sub


    Public Sub RecallTableState(check As String)
        Dim iTable As New zclsTable
        iTable.RecallTableState(check)
        iTable = Nothing
    End Sub

    Public Sub UnassignCheck(check As String)
        Dim iTable As New zclsTable
        iTable.Table = ThisTable.Table
        iTable.UnassignCheck(check)
        iTable = Nothing
    End Sub

    Public Function GetTableChecks(table As String) As Collection
        Dim iTable As New zclsTable(table)
        Return iTable.Checks
    End Function
End Module
