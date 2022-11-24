
Public MustInherit Class DBObject


    Public ReadOnly Property Rs As Recordset
        Get
            Return GetRs()
        End Get
    End Property

    Public ReadOnly Property Conn As Connection
        Get
            Return GetConn()
        End Get
    End Property

    Public ReadOnly Property Db As String
        Get
            Return GetDb()
        End Get
    End Property

    Public ReadOnly Property DbFile As String
        Get
            Return GetDbFile()
        End Get
    End Property

    Public MustOverride Function GetDb() As String

    Public MustOverride Function GetDbFile() As String

    Public MustOverride Function GetConn() As ADODB.Connection

    Public MustOverride Function GetRs() As ADODB.Recordset

    Public Sub OpenDbs()
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(iDataObject)
        iDataObject.OpenDbs(iDataObject)
        iDataObject = Nothing
    End Sub

    Public Sub CloseDbs()
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(iDataObject)
        iDataObject.CloseDbs(iDataObject)
        iDataObject = Nothing
    End Sub

    Public Function Wrap(obj As aclsDataObject) As aclsDataObject
        Dim iDataObj As New aclsDataObject
        iDataObj = obj
        iDataObj.Rs = Me.Rs
        iDataObj.Conn = Me.Conn
        iDataObj.Db = Me.Db
        iDataObj.DbFile = Me.DbFile
        Wrap = iDataObj
        iDataObj = Nothing
    End Function


End Class

'Public MustInherit Class DBObject


'    Public MustOverride ReadOnly Property Rs As ADODB.Recordset
'    Public MustOverride ReadOnly Property Conn As ADODB.Connection
'    Public MustOverride ReadOnly Property Db As String
'    Public MustOverride ReadOnly Property DbFile As String

'    Public Sub OpenDbs()
'        Dim iDataObject As New aclsDataObject
'        iDataObject = Wrap(iDataObject)
'        iDataObject.OpenDbs(iDataObject)
'        iDataObject = Nothing
'    End Sub

'    Public Sub CloseDbs()
'        Dim iDataObject As New aclsDataObject
'        iDataObject = Wrap(iDataObject)
'        iDataObject.CloseDbs(iDataObject)
'        iDataObject = Nothing
'    End Sub

'    Public Function Wrap(obj As aclsDataObject) As aclsDataObject
'        Dim iDataObj As New aclsDataObject
'        iDataObj = obj
'        iDataObj.Rs = Me.Rs
'        iDataObj.Conn = Me.Conn
'        iDataObj.Db = Me.Db
'        iDataObj.DbFile = Me.DbFile
'        Wrap = iDataObj
'        iDataObj = Nothing
'    End Function


'End Class
