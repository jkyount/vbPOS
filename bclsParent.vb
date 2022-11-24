Public Class bclsParent

    Private pID As Integer
    Private pItem As aclsItem
    Private pItemID As Integer

    Public Property ID As Integer
        Set(value As Integer)
            pID = value
        End Set
        Get
            Return pID
        End Get
    End Property

    Public Property item As aclsItem
        Set(value As aclsItem)
            pItem = value
        End Set
        Get
            Return GetItemByID(Me.ID)
        End Get
    End Property

    Public Property ItemID As Integer
        Set(value As Integer)
            pItemID = value
        End Set
        Get
            Return pItemID
        End Get
    End Property

    Public Sub New(ID As Integer)
        Me.ID = ID
    End Sub

    Public Sub New()

    End Sub

    'Public Function GetNew(ID As Integer) As bclsParent
    '    Dim x As New bclsParent
    '    x.ID = ID
    '    GetNew = x
    '    x = Nothing
    'End Function

    Public Sub SetParent()
        ThisItem.Parent = New bclsParent(-1)
    End Sub

    Public Sub InheritParentPrintParams(item As aclsItem)
    End Sub

    Public Sub BuildItemCollection()
        ClearCollection(CItem)
        ThisItem.IsPrimaryItem = True
        CItem.Add(ThisItem, CStr(ThisItem.CollID))
        SetCurrentParent(ThisItem.CollID)
    End Sub

    Public Sub RefreshPreviewWindow()
        RefreshPreviewWindow()
    End Sub

    Public Sub CheckForRequiredComponents(item As aclsItem)
        If item.ItemOptions = False Then
            ' TODO LINES BELOW NEED TO BE IMPLEMENTED
            'HideShape "frmSaladFrame"
            OrderQueuedItems(CItem)
            Exit Sub
        End If

        If Not item.RequiredComponents.Count = 0 Then
            If MissingComponents = True Then
                Exit Sub
            End If
        End If
        ' TODO LINES BELOW NEED TO BE IMPLEMENTED
        SetCurrentParent(item.CollID)
        '        HideShape "ComponentBLOCK"
        'HideShape "grpScrollCategoryItems"
        'ShowShape "btnDone"
        'SetCurrentFamily("")
    End Sub

End Class
