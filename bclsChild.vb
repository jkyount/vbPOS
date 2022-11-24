Public Class bclsChild
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

    Public Sub New()

    End Sub

    Public Sub New(ID As Integer)
        Me.ID = ID
    End Sub

    'Public Function GetNew(ID As Integer) As bclsChild
    '    Dim x As New bclsChild
    '    x.ID = ID
    '    GetNew = x
    '    x = Nothing
    'End Function

    Public Sub SetParent()
        ThisItem.Parent = New bclsParent(CurrentParent)
    End Sub

    Public Sub InheritParentPrintParams(item As aclsItem)
    End Sub

    Public Sub BuildItemCollection()
        If ThisItem.Family = "Sce" Then
            If CItem("1").AltScePrice = True Then
                ThisItem.Price = ThisItem.AltPrice
            End If
        End If
        GetItemByID(ThisItem.Parent.ID).Children.coll.Add(New bclsChild(ThisItem.CollID), CStr(ThisItem.CollID))
        ThisItem.IsPrimaryItem = False
        CItem.Add(ThisItem, CStr(ThisItem.CollID))
    End Sub

    Public Sub RefreshPreviewWindow()
        RefreshPreviewWindow()
    End Sub


    Public Sub CheckForRequiredComponents(item As aclsItem)
        SetCurrentParent(item.CollID)
        If Not item.RequiredComponents.Count = 0 Then
            If MissingComponents = True Then
                Exit Sub
            End If
        End If
        CItem("1").OrderRank.CheckForRequiredComponents(CItem("1"))
    End Sub

End Class
