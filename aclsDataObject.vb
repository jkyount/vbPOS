Public Class aclsDataObject

    Private pField1 As String
    Private pField2 As String
    Private pValue1 As Object
    Private pValue2 As Object
    Private pDb As String
    Private pDbFile As String
    Private pConn As New ADODB.Connection
    Private pRs As New ADODB.Recordset
    Private pArchive As String
    Private pArchiveDbFile As String

    Property Db() As String
        Get
            Return pDb
        End Get
        Set(val As String)
            pDb = val
        End Set
    End Property

    Property Field1() As String
        Get
            Return pField1
        End Get
        Set(val As String)
            pField1 = val
        End Set
    End Property
    Property Field2() As String
        Get
            Return pField2
        End Get
        Set(val As String)
            pField2 = val
        End Set
    End Property
    Property Value1() As Object
        Get
            Return pValue1
        End Get
        Set(val As Object)
            pValue1 = val
        End Set
    End Property
    Property Value2() As Object
        Get
            Return pValue2
        End Get
        Set(val As Object)
            pValue2 = val
        End Set
    End Property
    Property DbFile() As String
        Get
            Return pDbFile
        End Get
        Set(val As String)

            pDbFile = val
        End Set
    End Property

    Property Conn() As ADODB.Connection
        Get
            Return pConn
        End Get
        Set(val As ADODB.Connection)
            pConn = val
        End Set
    End Property


    Property Rs() As ADODB.Recordset
        Get
            Return pRs
        End Get
        Set(val As ADODB.Recordset)
            pRs = val
        End Set
    End Property
    Property Archive() As String
        Get
            Return pArchive
        End Get
        Set(val As String)
            pArchive = val
        End Set
    End Property
    Property ArchiveDbFile() As String
        Get
            Return pArchiveDbFile
        End Get
        Set(val As String)
            pArchiveDbFile = val
        End Set
    End Property

    Public Function GetNewMatchObj(Optional Where As String = "CheckNumber", Optional Equals As Object = Nothing, Optional AndWhere As String = "", Optional Equals2 As Object = Nothing) As aclsDataObject
        Dim x As New aclsDataObject
        x.QueryParams(Where, Equals, AndWhere, Equals2)
        Return x
        x = Nothing
    End Function

    Public Function GetNewUpdateObj(Optional Where As String = "CheckNumber", Optional Equals As Object = Nothing, Optional UpdateField As String = "", Optional UpdateValue As Object = Nothing) As aclsDataObject
        Dim x As New aclsDataObject
        x.QueryParams(Where, Equals, UpdateField, UpdateValue)
        Return x
        x = Nothing
    End Function

    Public Function GetNewArchiveObj(obj As Object, DataObj As aclsDataObject) As aclsDataObject
        DataObj.Db = obj.GetArchive
        DataObj.DbFile = obj.GetArchiveDbFile
        Return DataObj
    End Function


    Public Sub QueryParams(Optional a As String = "ItemID", Optional b As Object = Nothing, Optional y As String = "", Optional z As Object = Nothing)
        Me.Field1 = a
        If IsNothing(b) = True Then b = ""
        Me.Value1 = b
        Me.Field2 = y
        If IsNothing(z) = True Then z = ""
        Me.Value2 = z
    End Sub

    Public Sub OpenDbs(DataObject As aclsDataObject)
        '5/18 DataObject As Variant/DataObject As aclsDataObject
        Dim Conn As ADODB.Connection
        Conn = New ADODB.Connection
        Conn = DataObject.Conn
        Conn.CursorLocation = CursorLocationEnum.adUseClient
        Dim rs As New ADODB.Recordset
        rs = DataObject.Rs
        Conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Jared\POS\Access_VS\" & DataObject.DbFile & ".accdb;Persist Security Info=False;"
        Conn.Open()
        rs.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        rs.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rs.ActiveConnection = Conn
        rs = Nothing
        Conn = Nothing
    End Sub



    Public Sub CloseDbs(DataObject As aclsDataObject)
        '5/18 DataObject As Variant/DataObject As aclsDataObject
        Dim Conn As ADODB.Connection
        Conn = New ADODB.Connection
        Conn = DataObject.Conn
        Dim rs As New ADODB.Recordset
        rs = DataObject.Rs
        rs.Close()

        Conn.Close()
        rs = Nothing
        Conn = Nothing
    End Sub

    Public Function GetMatch(DataObject As aclsDataObject, MatchQuery As String) As ADODB.Recordset
        OpenDbs(DataObject)
        Dim rs As ADODB.Recordset
        rs = DataObject.Rs
        rs.let_Source(MatchQuery)
        On Error GoTo EH
        rs.Open()
        On Error GoTo 0
        GetMatch = rs
        Exit Function
EH:
        On Error GoTo 0
        rs.let_Source("SELECT * FROM " & DataObject.Db & " WHERE False")          '           " & rs.Fields(0).name & " = """""
        rs.Open()
        GetMatch = rs
    End Function

    Public Sub ArchiveData(DataObject As aclsDataObject, ArchiveCmd As String)
        OpenDbs(DataObject)
        DataObject.Conn.Execute(ArchiveCmd)
        DataObject.Conn.Close
    End Sub

End Class
