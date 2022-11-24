Imports Scripting

Public Class zclsTable
    Inherits DBObject
    Public dbTable As New ADODB.Connection
    Public rsTable As New ADODB.Recordset
    Private pTable As String
    Private pParentTable As String
    Private pServerName As String
    Private pHasOpenCheck As Boolean
    Private zclsOrder_pCheck As String
    Private pSplit As Boolean
    Private pChecks As New Collection
    Private pServerNum As Object
    Private pTableCheck As String



    Public Overrides Function GetDb() As String
        GetDb = "TableStates"
    End Function
    Public Overrides Function GetDbFile() As String
        GetDbFile = "TableStates"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        Return dbTable
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        Return rsTable
    End Function

    Public Property Checks() As Collection
        Get
            Return GetChecks()
        End Get
        Set(value As Collection)
            pChecks = value
        End Set
    End Property

    Public Property ServerNum() As Object
        Get
            Me.GetTableParam("ServerNum")
        End Get
        Set(value As Object)
            pServerNum = value
        End Set
    End Property

    Public Property Table() As String
        Get
            Return pTable
        End Get
        Set(value As String)
            pTable = value
        End Set
    End Property

    Public Property ParentTable() As String
        Get
            Return pParentTable
        End Get
        Set(value As String)
            pParentTable = value
        End Set
    End Property

    Public Property ServerName() As String
        Get
            Me.GetTableParam("ServerName")
        End Get
        Set(value As String)
            pServerName = value
        End Set
    End Property

    Public Property HasOpenCheck() As Boolean
        Get
            Me.GetTableParam("HasOpenCheck")
        End Get
        Set(value As Boolean)
            pHasOpenCheck = value
        End Set
    End Property



    Public Sub New()

    End Sub

    Public Sub New(ParentTable As String)
        Me.ParentTable = ParentTable
    End Sub



    '==========================================================================
    Public Sub Assign(check As String, ServerNum As Integer, ServerName As String)
        Me.Table = GetNextTable(Me.ParentTable)
        Me.AssignCheck(check)
        Me.AssignServer(ServerNum, ServerName)
    End Sub

    Public Function GetTableParam(param As String) As Object
        OpenDbs()
        rsTable.let_Source("SELECT * FROM TableStates WHERE Table = """ & Me.Table & """")
        rsTable.Open()
        GetTableParam = rsTable.Fields(param).Value
        rsTable.Close()
        dbTable.Close()
    End Function

    Public Sub SetTableInUse(check As String, state As String)
        pSetTableInUse(check, state)
    End Sub

    Private Sub pSetTableInUse(check As String, state As String)
        OpenDbs()
        rsTable.let_Source("SELECT * From TableStates WHERE CheckNumber = """ & check & """")
        rsTable.Open()
        rsTable.Fields("InUse").Value = state
        rsTable.Fields("CheckNumber").Value = check
        rsTable.Update()
        rsTable.Close()
        dbTable.Close()
    End Sub

    Public Function GetTablesInUse() As Collection
        Return pGetTablesInUse()
    End Function

    Private Function pGetTablesInUse() As Collection
        Dim coll As New Collection
        Dim x As String
        OpenDbs()
        rsTable.let_Source("SELECT * From TableStates WHERE ParentTable = True AND InUse = True ORDER BY Table")
        rsTable.Open()

        Do Until rsTable.EOF
            If Not rsTable.Fields("ParentTable").Value = x Then
                x = rsTable.Fields("ParentTable").Value
                coll.Add(x)
            End If
            rsTable.MoveNext()
        Loop
        pGetTablesInUse = coll
        rsTable.Close()
        dbTable.Close()
    End Function

    Public Function GetTableStates() As Collection
        Return pGetTableStates()
    End Function
    Private Function pGetTableStates() As Collection



        Dim iDataObject As aclsDataObject = Wrap(GetNewMatchObj("IsParentTable", True))

        Dim coll As Collection = CDictCollection(GetRecordsetMatch(iDataObject, ConstructOrderedQuery(iDataObject, " ORDER BY Table")))



        Return coll
    End Function



    Public Function GetServerTables(ServerNum As Integer) As Collection
        Return pGetServerTables(ServerNum)
    End Function
    Private Function pGetServerTables(ServerNum As Integer) As Collection
        Dim coll As New Collection
        Dim x As String
        OpenDbs()
        rsTable.let_Source("SELECT * From TableStates WHERE InUse = True AND ServerNum = " & ServerNum & " ORDER BY Table")
        rsTable.Open()

        Do Until rsTable.EOF
            If Not rsTable.Fields("ParentTable").Value = x Then
                x = rsTable.Fields("ParentTable").Value
                coll.Add(x)
            End If
            rsTable.MoveNext()
        Loop
        pGetServerTables = coll
        rsTable.Close()
        dbTable.Close()
    End Function


    Public Function GetNextTable(ParentTable As String) As String
        Return pGetNextTable(ParentTable)
    End Function
    Private Function pGetNextTable(ParentTable As String) As String
        OpenDbs()
        rsTable.let_Source("SELECT * FROM TableStates WHERE ParentTable = """ & ParentTable & """ ORDER BY LEN(Table), Table ASC")
        rsTable.Open()

        Do Until rsTable.Fields("InUse").Value = False Or rsTable.EOF
            rsTable.MoveNext()
        Loop

        If rsTable.EOF = True Then
            MsgBox("Could not assign check to table")
            Exit Function
        End If
        pGetNextTable = rsTable.Fields("Table").Value
        rsTable.Close()
        dbTable.Close()
    End Function

    Public Sub AssignCheck(check As String)
        pAssignCheck(check)
    End Sub
    Private Sub pAssignCheck(check As String)
        OpenDbs()
        rsTable.let_Source("SELECT * FROM TableStates WHERE Table = """ & Me.Table & """")
        rsTable.Open()
        rsTable.Fields("CheckNumber").Value = check
        rsTable.Fields("InUse").Value = True
        rsTable.Update()
        rsTable.Close()
        dbTable.Close()
    End Sub

    Public Sub UnassignCheck(check As String)
        pUnassignCheck(check)
    End Sub
    Private Sub pUnassignCheck(check As String)
        OpenDbs()
        rsTable.let_Source("SELECT * FROM TableStates WHERE CheckNumber = """ & check & """")
        rsTable.Open()
        rsTable.Fields("CheckNumber").Value = ""
        rsTable.Fields("InUse").Value = False
        rsTable.Fields("ServerName").Value = ""
        rsTable.Fields("ServerNum").Value = 0
        rsTable.Update()
        rsTable.Close()
        dbTable.Close()
    End Sub

    Public Sub AssignServer(ServerNum As Integer, ServerName As String)
        pAssignServer(ServerNum, ServerName)
    End Sub
    Private Sub pAssignServer(ServerNum As Integer, ServerName As String)
        OpenDbs()
        rsTable.let_Source("SELECT * FROM TableStates WHERE Table = """ & Me.Table & """")
        rsTable.Open()
        rsTable.Fields("ServerNum").Value = ServerNum
        rsTable.Fields("ServerName").Value = ServerName
        rsTable.Update()
        rsTable.Close()
        dbTable.Close()
    End Sub

    Private Function GetChecks() As Collection
        Dim coll As New Collection
        OpenDbs()
        rsTable.let_Source("SELECT * FROM TableStates WHERE ParentTable = """ & Me.ParentTable & """ AND NOT CheckNumber = """"")
        rsTable.Open()

        Do Until rsTable.EOF
            coll.Add(rsTable.Fields("CheckNumber").Value)
            rsTable.MoveNext()
        Loop
        rsTable.Close()
        dbTable.Close()
        GetChecks = coll
        coll = Nothing
    End Function

    Public Sub ClearTableStates()
        pClearTableStates()
    End Sub
    Private Sub pClearTableStates()
        OpenDbs()
        rsTable.let_Source("TableStates")
        rsTable.Open()

        Do Until rsTable.EOF
            rsTable.Fields("ServerName").Value = ""
            rsTable.Fields("CheckNumber").Value = ""
            rsTable.Fields("ServerNum").Value = 0
            rsTable.Fields("InUse").Value = False
            rsTable.MoveNext()
        Loop
        rsTable.UpdateBatch()
        CloseDbs()
    End Sub
    Public Sub RecallTableState(check As String)
        Dim dict As New Dictionary(Of String, Object)
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(GetNewMatchObj("CheckNumber", check))
        dict = CDictCollection(GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject)))(1)

        SetTableState(dict("ParentTable"), dict("Table"))
        dict = Nothing
        iDataObject = Nothing

    End Sub
End Class
