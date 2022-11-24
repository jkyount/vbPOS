Imports Scripting

Public Class zclsFamily
    Inherits DBObject

    Public dbFamily As New ADODB.Connection
    Public rsFamily As New ADODB.Recordset
    Private pStyleDict As Dictionary(Of String, Object)
    Private pCount As Integer
    Private pFamily As String
    Private pMembers() As Object
    Private pMenuStyle As Integer
    Private pFamilyGroup As String
    Private pFamilyID As Integer
    Private pMenuStyleObj As aclsMenuStyle


    Public ReadOnly Property IsMultiMenu As Boolean
        Get
            If Me.FamilyID = 0 Then
                Return False
                Exit Property
            End If
            Return ValueMatch(Me.Wrap(GetNewMatchObj("ID", Me.FamilyID)), "MultiMenu")
        End Get
    End Property

    Public Property FamilyID As Integer
        Set(value As Integer)
            pFamilyID = value
        End Set
        Get
            Return pFamilyID
        End Get
    End Property

    Public Property Count As Integer
        Set(value As Integer)
            pCount = value
        End Set
        Get
            Return GetCount(Me)
        End Get
    End Property

    Public Property Family As String
        Set(value As String)
            pFamily = value
        End Set
        Get
            Return GetFamily(Me.FamilyID)
        End Get
    End Property

    Public Property FamilyGroup As String
        Set(value As String)
            pFamilyGroup = value
        End Set
        Get
            Return GetFamilyGroup()
        End Get
    End Property

    Public Property Members As Object
        Set(value As Object)
            pMembers = value
        End Set
        Get
            Return GetMembers()
        End Get
    End Property

    Public Property MenuStyle As Integer
        Set(value As Integer)
            pMenuStyle = value
        End Set
        Get
            Return GetMenuStyle(Me)
        End Get
    End Property

    Public Property MenuStyleObj As aclsMenuStyle
        Set(value As aclsMenuStyle)
            pMenuStyleObj = value
        End Set
        Get
            Return GetMenuStyleObj(Me)
        End Get
    End Property

    Public Overrides Function GetDb() As String
        GetDb = "Family"
    End Function

    Public Overrides Function GetDbFile() As String
        GetDbFile = "Menu"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbFamily
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsFamily
    End Function


    Public Sub New()

    End Sub

    Public Sub New(FamilyID As Integer)
        Me.FamilyID = FamilyID
    End Sub

    Public Sub New(FamilyName As String)
        Dim FamilyID As Integer
        FamilyID = ValueMatch(Wrap(GetNewMatchObj("Family", FamilyName)), "ID")
        Me.FamilyID = FamilyID
    End Sub


    Private Function GetMembers() As Object
        Dim qry As String
        Dim iMenu As New zclsMenu
        Dim iDataObj As aclsDataObject
        If ValueMatch(Wrap(GetNewMatchObj("Family", Me.Family)), "MultiMenu") = True Then
            iDataObj = Wrap(GetNewMatchObj)
            qry = "SELECT Family FROM Family WHERE MultiMenuParent = """ & Me.Family & """"
            GetMembers = RsToArray(GetRecordsetMatch(iDataObj, qry))
            iDataObj.CloseDbs(iDataObj)
            iMenu = Nothing
            iDataObj = Nothing
            Exit Function
        End If
        iDataObj = iMenu.Wrap(GetNewMatchObj)
        qry = "SELECT ID FROM AllItems WHERE Family = """ & Me.Family & """ AND NOT ItemName = """" ORDER BY LEN(ID), ID ASC"
        GetMembers = RsToArray(GetRecordsetMatch(iDataObj, qry))
        iDataObj.CloseDbs(iDataObj)
        iMenu = Nothing
    End Function

    Public Function GetMembersColl() As Collection
        Dim qry As String
        Dim iMenu As New zclsMenu
        qry = "SELECT ID, ItemName FROM AllItems WHERE Family = """ & Me.Family & """ AND NOT ItemName = """" ORDER BY LEN(ID), ID ASC"
        Dim rs As ADODB.Recordset
        rs = GetRecordsetMatch(iMenu.Wrap(GetNewMatchObj), qry)
        Dim coll As New Collection
        Dim dict As Dictionary(Of String, Object)
        Dim i As Integer
        Dim cyclecount As Integer
        Dim ID As Integer
        Dim ItemName As String
        cyclecount = 1
        i = 1
        Dim BtnsPerPage As Integer
        'TODO -- UnDummy BtnsPerPage
        BtnsPerPage = 24
        'BtnsPerPage = Me.StyleDict("BtnsPerPage")
        dict = New Dictionary(Of String, Object)
        For i = 1 To rs.RecordCount
            If i = (cyclecount * BtnsPerPage) + 1 Then
                coll.Add(dict)
                dict = New Dictionary(Of String, Object)
                cyclecount = cyclecount + 1
            End If
            ID = rs.Fields("ID").Value
            ItemName = rs.Fields("ItemName").Value
            dict.Add(ID, ItemName)
            rs.MoveNext()
        Next i
        coll.Add(dict)
        GetMembersColl = coll
        iMenu.CloseDbs
        iMenu = Nothing
        coll = Nothing
        dict = Nothing
        rs = Nothing
        iMenu = Nothing
    End Function
    Private Function GetFamily(FamilyID As Integer)
        GetFamily = ValueMatch(Wrap(GetNewMatchObj("ID", FamilyID)), "Family")
    End Function
    Private Function GetFamilyGroup() As String
        GetFamilyGroup = ValueMatch(Wrap(GetNewMatchObj("Family", Me.Family)), "FamilyGroup")
    End Function

    Private Function GetMenuStyle(iFamily As zclsFamily) As Integer
        GetMenuStyle = ValueMatch(Wrap(GetNewMatchObj("Family", iFamily.Family)), "MenuStyle")
    End Function

    Private Function GetMenuStyleObj(iFamily As zclsFamily) As aclsMenuStyle
        Return New aclsMenuStyle(iFamily.MenuStyle)
    End Function



    Public Function GetCount(iFamily As zclsFamily) As Integer
        If iFamily.Family = "" Then
            GetCount = 0
        End If
        Dim iMenu As New zclsMenu
        GetCount = CountMatch(iMenu.Wrap(GetNewMatchObj("Family", iFamily.Family, "NOT ItemName", "")))
        iMenu = Nothing
    End Function

    '    Public Function GetStyleDict(iFamily As zclsFamily) As Dictionary
    '        Dim iMenuStyle As New aclsMenuStyle
    'Set iMenuStyle = iMenuStyle.GetNewMenuStyleObj(iFamily.MenuStyle)
    'Set GetStyleDict = iMenuStyle.StyleDict
    'Set iMenuStyle = Nothing
    'End Function

    'Public Sub AddNew(FamilyDict As Dictionary)
    '    Dim NewFamily As String
    '    NewFamily = FamilyDict("Family")
    '    Dim iDataObj As New aclsDataObject
    '    iDataObj = Wrap(GetNewMatchObj("Family", NewFamily))
    '    iDataObj.OpenDbs(iDataObj)
    '    iDataObj.Rs.Source = iDataObj.Db
    '    iDataObj.Rs.Open()
    '    iDataObj.Rs.AddNew(Array("Family"), Array(NewFamily))
    '    iDataObj.Rs.Update()
    '    iDataObj.CloseDbs(iDataObj)
    '    UpdateFromDict(iDataObj, FamilyDict)
    'End Sub

    Public Sub Remove(Family As String)

        Me.Family = Family

        DeleteMatch(Wrap(GetNewMatchObj("Family", Family)))
        Dim iMenu As New zclsMenu
        Update(iMenu.Wrap(GetNewUpdateObj("Family", Family, "Family", "Unassigned")))
        iMenu = Nothing

    End Sub

    Public Sub RemoveFromGUI(iFamily As zclsFamily)
        Update(Wrap(GetNewUpdateObj("Family", iFamily.Family, "Active", False)))

    End Sub

    Public Sub Activate()
        Update(Wrap(GetNewUpdateObj("Family", Me.Family, "Active", True)))
    End Sub

    Public Sub Deactivate()
        Update(Wrap(GetNewUpdateObj("Family", Me.Family, "Active", False)))
    End Sub

End Class
