Public MustInherit Class BaseItemClass
    Private pItemID As Integer

    Public Property ItemID As Integer
        Set(value As Integer)
            pItemID = value
        End Set
        Get
            Return pItemID
        End Get
    End Property


    Public Sub ItemInitialize(ItemID As Integer)
        Me.ItemID = ItemID
    End Sub

    Public Overridable Sub SpclConfig(item As aclsItem)
    End Sub

    Public Overridable Sub InheritParentPrintParams(item As aclsItem)
    End Sub

    Public Overridable Sub zclsItem_RefreshPreviewWindow()
    End Sub

    Public Overridable Sub zclsItem_UpdateGUI()
    End Sub

    Public Overridable Function GetItemIndicator() As String
        Return ""
    End Function

End Class

Public Class zclsPrimary
    Inherits BaseItemClass

    Public Overrides Sub zclsItem_RefreshPreviewWindow()
        'RefreshPreviewWindow
    End Sub

    Public Overrides Sub zclsItem_UpdateGUI()
        'ShowShape "frmSaladFrame"
    End Sub


    Public Overrides Function GetItemIndicator() As String
        GetItemIndicator = "> "
    End Function

End Class

Public Class zclsComponent
    Inherits BaseItemClass

    Public Overrides Sub zclsItem_RefreshPreviewWindow()
        'RefreshPreviewWindow
    End Sub

    Public Overrides Sub zclsItem_UpdateGUI()
        'HideShape "grpgui" & ThisItem.Family
    End Sub

    Public Overrides Function GetItemIndicator() As String
        GetItemIndicator = ""
    End Function
End Class

'Public Class zclsPrimary
'    Inherits BaseItemClass
'    Private pItemID As Integer

'    Public Property ItemID As Integer
'        Set(value As Integer)
'            pItemID = value
'        End Set
'        Get
'            Return pItemID
'        End Get
'    End Property


'    Public Sub ItemInitialize(ItemID As Integer)
'        Me.ItemID = ItemID
'        'ThisItem.ItemID = ItemID
'    End Sub

'    Public Sub SpclConfig(item As aclsItem)

'    End Sub

'    Public Sub InheritParentPrintParams(item As aclsItem)
'    End Sub


'    Public Sub zclsItem_RefreshPreviewWindow()
'        'RefreshPreviewWindow
'    End Sub

'    Public Sub zclsItem_UpdateGUI()
'        'ShowShape "frmSaladFrame"
'    End Sub


'    Public Function GetItemIndicator() As String
'        GetItemIndicator = "> "
'    End Function

'End Class