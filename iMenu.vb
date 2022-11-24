Module iMenu

    Public Function GetMenuObj() As zclsMenu
        Dim iMenu As zclsMenu
        iMenu = New zclsMenu
        GetMenuObj = iMenu
        iMenu = Nothing
    End Function
    Public Function GetNextItemID(Optional Family As String = "") As String
        'STILL STRING
        Dim iMenu As New zclsMenu
        GetNextItemID = iMenu.GetNextItemID(Family)
        iMenu = Nothing
    End Function

    Public Sub CreateNewSpecialItem(ItemID As Integer, ItemName As String, Price As Double)
        Dim iMenu As New zclsMenu
        iMenu.CreateNewSpecialItem(ItemID, ItemName, Price)
        iMenu = Nothing
    End Sub

    Public Function GetItemClassCode(ItemID As Integer) As Integer
        Dim iMenu As New zclsMenu
        GetItemClassCode = iMenu.GetItemClassCode(ItemID)

    End Function

End Module
