Module iItemClass
    Public Function GetItemType(ClassCode As String) As BaseItemClass
        Dim iItemclass As New zclsItemClass
        Return iItemclass.GetItemType(ClassCode)
    End Function
End Module
