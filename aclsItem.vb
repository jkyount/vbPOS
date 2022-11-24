Imports Scripting
Imports System.Security.Policy

Public Class aclsItem
    Private pItemID As Integer
    Private pItemName As String
    Private pPrice As Double
    Private pFamily As String
    Private pReq1 As String
    Private pReq2 As String
    Private pIsPrimaryItem As Boolean
    Private pCategory As String
    Private pCoolerFlag As Boolean
    Private pAlwaysTax As Boolean
    Private pAltScePrice As Boolean
    Private pAltPrice As Double
    Private pDiscountPrice As Double
    Private pPrintKitchen As Boolean
    Private pPrintPantry As Boolean
    Private pItemOptions As Boolean
    Private pClassCode As Integer
    Private pSeat As Integer
    Private pItemIndicator As String
    Private pItemType As BaseItemClass
    Private pParentID As Integer
    Private pChildID As Integer
    Private pCollID As Integer
    Private pComponents As Collection
    Private pParent As bclsParent
    Private pChild As bclsChild
    Private pChildren As bclsChildren
    Private pRequiredComponents As Dictionary(Of String, Object)
    Private pOrderRank As Object



    Public Property ItemID As Integer
        Set(value As Integer)
            pItemID = value
        End Set
        Get
            Return pItemID
        End Get
    End Property

    Public Property ItemName As String
        Set(value As String)
            pItemName = value
        End Set
        Get
            Return pItemName
        End Get
    End Property

    Public Property Price As Double
        Set(value As Double)
            pPrice = value
        End Set
        Get
            Return pPrice
        End Get
    End Property

    Public Property Family As String
        Set(value As String)
            pFamily = value
        End Set
        Get
            Return pFamily
        End Get
    End Property


    Public Property Req1 As String
        Set(value As String)
            pReq1 = value
        End Set
        Get
            Return pReq1
        End Get
    End Property

    Public Property Req2 As String
        Set(value As String)
            pReq2 = value
        End Set
        Get
            Return pReq2
        End Get
    End Property

    Public Property IsPrimaryItem As Boolean
        Set(value As Boolean)
            pIsPrimaryItem = value
        End Set
        Get
            Return pIsPrimaryItem
        End Get
    End Property

    Public Property Category As String
        Set(value As String)
            pCategory = value
        End Set
        Get
            Return pCategory
        End Get
    End Property

    Public Property CoolerFlag As Boolean
        Set(value As Boolean)
            pCoolerFlag = value
        End Set
        Get
            Return pCoolerFlag
        End Get
    End Property

    Public Property AlwaysTax As Boolean
        Set(value As Boolean)
            pAlwaysTax = value
        End Set
        Get
            Return pAlwaysTax
        End Get
    End Property

    Public Property AltScePrice As Boolean
        Set(value As Boolean)
            pAltScePrice = value
        End Set
        Get
            Return pAltScePrice
        End Get
    End Property

    Public Property AltPrice As Double
        Set(value As Double)
            pAltPrice = value
        End Set
        Get
            Return pAltPrice
        End Get
    End Property

    Public Property DiscountPrice As Double
        Set(value As Double)
            pDiscountPrice = value
        End Set
        Get
            Return pDiscountPrice
        End Get
    End Property

    Public Property ItemOptions As Boolean
        Set(value As Boolean)
            pItemOptions = value
        End Set
        Get
            Return pItemOptions
        End Get
    End Property

    Public Property PrintKitchen As Boolean
        Set(value As Boolean)
            pPrintKitchen = value
        End Set
        Get
            Return pPrintKitchen
        End Get
    End Property

    Public Property PrintPantry As Boolean
        Set(value As Boolean)
            pPrintPantry = value
        End Set
        Get
            Return pPrintPantry
        End Get
    End Property

    Public Property ClassCode As Integer
        Set(value As Integer)
            pClassCode = value
        End Set
        Get
            Return pClassCode
        End Get
    End Property

    Public Property seat As Integer
        Set(value As Integer)
            pSeat = value
        End Set
        Get
            Return pSeat
        End Get
    End Property

    Public Property ItemIndicator As String
        Set(value As String)
            pItemIndicator = value
        End Set
        Get
            Return Me.ItemType.GetItemIndicator
        End Get
    End Property

    Public Property ItemType As BaseItemClass
        Set(value As BaseItemClass)
            pItemType = value
        End Set
        Get
            Return pItemType
        End Get
    End Property

    Public Property ParentID As Integer
        Set(value As Integer)
            pParentID = value
        End Set
        Get
            Return pParentID
        End Get
    End Property

    Public Property CollID As Integer
        Set(value As Integer)
            pCollID = value
        End Set
        Get
            Return pCollID
        End Get
    End Property

    Public Property ChildID As Integer
        Set(value As Integer)
            pChildID = value
        End Set
        Get
            Return pChildID
        End Get
    End Property

    Public Property Components As Collection
        Set(value As Collection)
            pComponents = value
        End Set
        Get
            Return pComponents
        End Get
    End Property

    Public Property Parent As bclsParent
        Set(value As bclsParent)
            pParent = value
        End Set
        Get
            Return pParent
        End Get
    End Property

    Public Property Children As bclsChildren
        Set(value As bclsChildren)
            pChildren = value
        End Set
        Get
            Return pChildren
        End Get
    End Property

    Public Property RequiredComponents As Dictionary(Of String, Object)
        Set(value As Dictionary(Of String, Object))
            pRequiredComponents = value
        End Set
        Get
            Return pRequiredComponents
        End Get
    End Property

    Public Property OrderRank As Object
        Set(value As Object)
            pOrderRank = value
        End Set
        Get
            Return pOrderRank
        End Get
    End Property

    Public Sub Initialize(ItemID As Integer)
        Dim ItemType As BaseItemClass
        ItemType = GetItemType(CStr(GetItemClassCode(ItemID)))
        ItemType.ItemInitialize(ItemID)



        ThisItem.ItemType = ItemType
        ThisItem = ThisItem.CreateNew(ThisItem.ItemType.ItemID)
        ItemType = Nothing
    End Sub
    Public Function CreateNew(ItemID As Integer) As aclsItem
        Dim item As New aclsItem
        item = PopItem(ItemID, item)
        item.ItemType = ThisItem.ItemType
        CreateNew = item
        item = Nothing
    End Function

    Public Function GetRequiredComponents() As Dictionary(Of String, Object)
        Dim dict As New Dictionary(Of String, Object)
        If Not Me.Req1 = "" Then
            dict.Add(Me.Req1, Me.Req1)
        End If
        If Not Me.Req2 = "" Then
            dict.Add(Me.Req2, Me.Req2)
        End If
        GetRequiredComponents = dict
    End Function

    Public Function ValidateRequirements(item As aclsItem, Requirement As String) As Boolean
        Dim member As aclsItem
        For Each member In CItem
            If member.Parent.ID = item.CollID Then
                If member.Family = Requirement Then
                    ValidateRequirements = True
                    Exit Function
                End If
            End If
        Next member
        ValidateRequirements = False
    End Function

    Public Function UnassignChild(ID As Integer)
        Dim child As bclsChild
        For Each child In Me.Children.coll
            If child.ID = ID Then
                Me.Children.coll.Remove(CStr(ID))
            End If
        Next child
    End Function

    Public Function AssignOrderRank() As Object
        If CItem.Count > 0 Then
            Dim iChild As New bclsChild
            AssignOrderRank = iChild
            iChild = Nothing
            Exit Function
        End If

        Dim iParent As New bclsParent
        AssignOrderRank = iParent
        iParent = Nothing
    End Function
End Class
