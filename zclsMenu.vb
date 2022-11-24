Public Class zclsMenu
    Inherits DBObject
    Public dbMenu As New ADODB.Connection
    Public rsMenu As New ADODB.Recordset

#Region "Overrides"

    Public Overrides Function GetDb() As String
        Return "AllItems"
    End Function

    Public Overrides Function GetDbFile() As String
        Return "Menu"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        Return dbMenu
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        Return rsMenu
    End Function

#End Region

    Public Function GetItemClassCode(ItemID As Integer) As Integer
        GetItemClassCode = pGetItemClassCode(ItemID)
    End Function

    Private Function pGetItemClassCode(ItemID As Integer) As Integer
        OpenDbs()
        rsMenu.let_Source("SELECT * FROM AllItems WHERE ID = " & ItemID & "")
        rsMenu.Open()
        pGetItemClassCode = rsMenu.Fields("ClassCode").Value
        rsMenu.Close()
        dbMenu.Close()
    End Function

    Public Function GetNextItemID(Family As String) As Integer
        'STILL STRING
        GetNextItemID = pGetNextItemID(Family)
    End Function

    Private Function pGetNextItemID(Optional Family As String = "") As Integer
        'STILL STRING
        Dim FamilyClause As String
        FamilyClause = ""
        If Not Family = "" Then
            FamilyClause = "Family = """ & Family & """ AND "
        End If
        OpenDbs()
        rsMenu.let_Source("SELECT * FROM AllItems WHERE " & FamilyClause & "ItemName = """" ORDER BY ID ASC")
        rsMenu.Open()

        If rsMenu.EOF Then
            rsMenu.AddNew()
            rsMenu.Fields("ItemID").Value = "Item" & rsMenu.Fields("ID").Value
            rsMenu.Update()
            pGetNextItemID = rsMenu.Fields("ID").Value
            rsMenu.Close()
            dbMenu.Close()
            Exit Function
        End If
        rsMenu.MoveFirst()
        pGetNextItemID = rsMenu.Fields("ID").Value
        rsMenu.Close()
        dbMenu.Close()
    End Function

    'Private Function pGetNextItemID(Family As String) As Integer
    ''STILL STRING
    'OpenDbs
    'rsMenu.Source = "SELECT * FROM AllItems WHERE Family = """ & Family & """ AND ItemName = """""
    'rsMenu.Open
    'rsMenu.MoveFirst
    'Dim pItemID As String
    'pItemID = rsMenu.Fields("ID").value
    'pGetNextItemID = pItemID
    'rsMenu.Close
    'dbMenu.Close
    'End Function

    Public Sub CreateNewSpecialItem(ItemID As Integer, ItemName As String, Price As Decimal)
        pCreateNewSpecialItem(ItemID, ItemName, Price)
    End Sub
    Private Sub pCreateNewSpecialItem(ItemID As Integer, ItemName As String, Price As Decimal)
        OpenDbs()
        rsMenu.let_Source("SELECT * FROM AllItems WHERE ID = " & ItemID & "")
        rsMenu.Open()
        rsMenu.Fields("ItemName").Value = ItemName
        rsMenu.Fields("Price").Value = Price
        rsMenu.Fields("Category").Value = "Food"
        rsMenu.Update()
        rsMenu.Close()
        dbMenu.Close()
    End Sub

    Public Sub ClearCustomItems()
        ClearNameAndPrice("CustomItem")
        ClearNameAndPrice("SpclInstruction")
    End Sub

    Private Sub ClearNameAndPrice(Family As String)
        Dim iMenu As New zclsMenu
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(GetNewMatchObj("Family", Family))
        rsMenu = GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject))
        If Not rsMenu.EOF Then
            Do Until rsMenu.EOF
                rsMenu.Fields("ItemName").Value = ""
                rsMenu.Fields("Price").Value = 0
                rsMenu.MoveNext()
            Loop
            rsMenu.UpdateBatch()
        End If
        CloseDbs()
    End Sub

    Public Function GetItemsInFamily(Family As String) As Object
        GetItemsInFamily = pGetItemsInFamily(Family)
    End Function

    Private Function pGetItemsInFamily(Family As String) As Object
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(iDataObject.GetNewMatchObj("Family", Family, "NOT ItemName", ""))
        Dim arr As Object
        arr = GetMatch(iDataObject)
        If arr = Nothing Then
            ReDim arr(0 To 0)
            Dim temparr(0 To 0, 0 To 0) As Object
            temparr(0, 0) = ""
            arr(0) = temparr
            pGetItemsInFamily = arr
            iDataObject = Nothing
            Exit Function
        End If
        pGetItemsInFamily = arr
        iDataObject = Nothing
    End Function

    Public Function PairItemNameAndID(arr As Object) As Object
        If UBound(arr(1)) < 2 Then
            PairItemNameAndID = arr
            Exit Function
        End If
        Dim i As Integer
        Dim TempArray As Object
        TempArray = arr
        For i = 1 To UBound(TempArray)
            TempArray(i)(0, 0) = TempArray(i)(0, 0) & "  -  " & TempArray(i)(2, 0)
        Next i
        PairItemNameAndID = TempArray

    End Function

    Public Sub Remove(ItemID As Integer)
        pRemove(ItemID)
    End Sub

    Private Sub pRemove(ItemID As Integer)
        Dim iMenu As New zclsMenu
        Dim rs As ADODB.Recordset
        rs = GetRecordsetMatch(Wrap(GetNewMatchObj("ID", ItemID)))
        Dim fld As Object
        For Each fld In rs.Fields
            If Not fld.name = "ID" Then
                If Not fld.name = "ItemID" Then
                    fld.value = False
                End If
            End If
            If fld.Type = DataTypeEnum.adVarWChar Then
                If Not fld.name = "ItemID" Then
                    fld.value = ""
                End If
            End If

        Next fld
        rs.Update()
        CloseDbs()
        iMenu = Nothing
    End Sub

End Class
