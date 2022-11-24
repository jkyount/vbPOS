Module guiOrder
    Public Function GetFamilyGroup(FamilyGroup As String) As Collection

        Dim iFamily As New zclsFamily
        Dim qry As String
        qry = "SELECT * FROM Family WHERE FamilyGroup = """ & FamilyGroup & """ AND Active = True AND MultiMenuMember = False ORDER BY ID ASC"
        GetFamilyGroup = CDictCollection(GetRecordsetMatch(iFamily.Wrap(GetNewMatchObj), qry))
        iFamily.CloseDbs()
        iFamily = Nothing
    End Function

    Public Function GetFamilyMembers(FamilyID As Integer) As Collection

        Dim iFamily As New zclsFamily
        iFamily.FamilyID = FamilyID

        GetFamilyMembers = iFamily.GetMembersColl

        iFamily = Nothing
    End Function

    Public Function FormatChildSpacing(coll As Collection) As Collection
        Dim i As Integer, k As Integer
        Dim SpacerCount As Integer
        Dim Spacer As String
        For i = 2 To coll.Count
            SpacerCount = 1
            Spacer = ""
            Dim item As aclsItem
            item = coll(i)
            Dim ParentID As Integer
            ParentID = item.Parent.ID
            Do Until ParentID = 1
                SpacerCount = SpacerCount + 1
                item = GetItemByID(item.Parent.ID)
                ParentID = item.Parent.ID
            Loop
            For k = 1 To SpacerCount
                Spacer = Spacer & "    "
            Next k
            coll(i).ItemName = Spacer & coll(i).ItemName
        Next i
        FormatChildSpacing = coll
    End Function
End Module
