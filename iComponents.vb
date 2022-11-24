Module iComponents
    'Public Function GetNewChild(ID As Integer) As bclsChild
    '    Dim iChild As New bclsChild
    '    GetNewChild = iChild.GetNew(ID)
    '    iChild = Nothing
    'End Function

    'Public Function GetNewChildren() As bclsChildren
    '    Dim iChildren As New bclsChildren
    '    GetNewChildren = iChildren.GetNew
    '    iChildren = Nothing
    'End Function

    '    Public Function GetNewParent(ID As Integer) As bclsParent
    '        Dim iParent As New bclsParent
    'Set GetNewParent = iParent.GetNew(ID)
    'Set iParent = Nothing
    'End Function

    Public Function GetItemByID(ID As Integer) As aclsItem
        Dim iItem As aclsItem
        For Each iItem In CItem
            If iItem.CollID = ID Then
                Return iItem
                Exit Function
            End If
        Next iItem

    End Function

    Public Function GetItemByParentID(ParentID As Integer) As aclsItem
        Dim iItem As aclsItem
        For Each iItem In CItem
            If iItem.Parent.ID = ParentID Then
                Return iItem
                Exit Function
            End If
        Next iItem
    End Function

    Public Function GetPrimaryItem() As aclsItem
        Dim iItem As aclsItem
        For Each iItem In CItem
            If iItem.Parent.ID = -1 Then
                Return iItem
                Exit Function
            End If
        Next iItem
    End Function

    Public Sub RemoveItemFromQueue(item As aclsItem)
        Dim child As bclsChild
        If item.Children.coll.Count > 0 Then
            For Each child In item.Children.coll
                RemoveItemFromQueue(GetItemByID(child.ID))
                item.Children.coll.Remove(CStr(child.ID))
            Next child
        End If
        CItem.Remove(CStr(item.CollID))

    End Sub

    Public Function NormalizePrintParameters(coll As Collection) As Collection

        Dim ThisChild As aclsItem

        For Each child As bclsChild In coll(1).Children.coll
            ThisChild = GetItemByID(child.ID)
            ThisChild.ItemType.InheritParentPrintParams(ThisChild)
            ApplyParentParamToChildren(ThisChild, ThisChild.PrintKitchen)
        Next child
        Return coll
    End Function

    Public Sub ApplyParentParamToChildren(item As aclsItem, param As Object)
        Dim child As bclsChild
        If item.Children.coll.Count > 0 Then
            For Each child In item.Children.coll
                ApplyParentParamToChildren(GetItemByID(child.ID), param)
                GetItemByID(child.ID).PrintKitchen = param
            Next child
        End If
        item.PrintKitchen = param
    End Sub

    Public Function FormatItemCollection(coll As Collection) As Collection
        Dim tempcoll As Collection
        tempcoll = FormatSides(FormatChildSpacing(OrderByParent(NormalizePrintParameters(coll))))

        'FormatChildSpacing tempcoll
        'FormatSides tempcoll
        Return tempcoll
    End Function


End Module
