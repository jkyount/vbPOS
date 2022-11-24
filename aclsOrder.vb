Imports Scripting

Public Class aclsOrder
    Private pCheck As String
    'Private pValueDict As Scripting.Dictionary
    Private pValueDict As Dictionary(Of String, Object)
    Private pOrderType As Object
    Private pOrderDetails As Dictionary(Of String, Object)
    Private pNewOrderDetails As Dictionary(Of String, Object)
    Private pTableVal As Boolean
    Private pTestProp As String

    Public Property TestProp() As String
        Set(value As String)
            pTestProp = value
        End Set
        Get
            Return pTestProp
        End Get
    End Property

    Public Property Check() As String
        Set(value As String)
            pCheck = value
        End Set
        Get
            Return pCheck
        End Get
    End Property

    Public Property ValueDict() As Dictionary(Of String, Object)
        Set(value As Dictionary(Of String, Object))
            pValueDict = value
        End Set
        Get
            Return pValueDict
        End Get
    End Property

    'Public Property ValueDict() As Scripting.Dictionary
    '    Set(value As Scripting.Dictionary)
    '        pValueDict = value
    '    End Set
    '    Get
    '        Return pValueDict
    '    End Get
    'End Property

    Public Property OrderDetails() As Dictionary(Of String, Object)
        Set(value As Dictionary(Of String, Object))
            pOrderDetails = value
        End Set
        Get
            Return pOrderDetails
        End Get
    End Property

    Public Property NewOrderDetails() As Dictionary(Of String, Object)
        Set(value As Dictionary(Of String, Object))
            pNewOrderDetails = value
        End Set
        Get
            Return NewOrderInfo(Me)
        End Get
    End Property


    Public Property OrderType() As Object
        Set(value As Object)
            pOrderType = value
        End Set
        Get
            Return pOrderType
        End Get
    End Property


    Public ReadOnly Property TableVal() As Boolean
        Get
            Return GetTableVal()
        End Get
    End Property



    Public ReadOnly Property Payments() As Decimal
        Get
            Return Me.ValueDict("Cash") + Me.ValueDict("Charge") + Me.ValueDict("GiftCert")
        End Get
    End Property


    Public ReadOnly Property AdjustedTotal() As Decimal
        Get
            Return Me.ValueDict("Total") - (Me.ValueDict("Cash") + Me.ValueDict("Charge") + Me.ValueDict("GiftCert"))
        End Get
    End Property


    Public Function NewOrderObject(OrderType As Object, check As String) As aclsOrder
        Dim iOrder As New aclsOrder
        iOrder.OrderType = OrderType
        iOrder.check = check
        iOrder.ImportCheckDetails(iOrder, check)
        NewOrderObject = iOrder
    End Function

    Public Sub SetOrderInfo(Order As aclsOrder, Optional dict As Dictionary(Of String, Object) = Nothing)
        If dict Is Nothing Then
            Order.OrderDetails = NewOrderInfo(Order)
        End If
        If dict IsNot Nothing Then
            Order.OrderDetails = dict
        End If
    End Sub

    Public Function NewOrderInfo(Order As aclsOrder) As Dictionary(Of String, Object)
        Dim dict As New Dictionary(Of String, Object) From {
            {"OrderName", Order.OrderType.GetOrderName},
            {"PickupTime", Order.OrderType.GetPickupTime},
            {"DineIn", Order.OrderType.GetDineIn},
            {"Table", Order.OrderType.GetTable},
            {"Phone", Order.OrderType.GetPhone},
            {"ServerNum", ThisEmployee.ServerNum},
            {"ServerName", ThisEmployee.FirstName},
            {"CheckDate", Format(Now, "MMddyy")},
            {"CheckOpen", Format(Now, "HH:mm")}
        }
        NewOrderInfo = dict
        dict = Nothing
    End Function

    Public Function SameOrderInfo() As Dictionary(Of String, Object)
        SameOrderInfo = Me.ValueDict
    End Function

    Public Function SameOrderType() As Object
        If Me.ValueDict("DineIn") = True Then
            Return New oclsDineIn
        End If
        If Me.ValueDict("DineIn") = False Then
            Return New oclsCarryout
        End If

    End Function

    'Public Sub ImportCheckDetails(OrderObj As aclsOrder, check As String)
    '    OrderObj.ValueDict = GetCheckAttributes(check)
    'End Sub

    Public Sub New()

    End Sub
    Public Sub New(check As String)
        Me.Check = check
        ImportCheckDetails(Me, check)
    End Sub
    Public Sub ImportCheckDetails(OrderObj As aclsOrder, check As String)
        OrderObj.ValueDict = GetCheckAttributes(check)
    End Sub

    Public Function SplitCheckInitialOrderDetails() As Dictionary(Of String, Object)
        Dim dict As New Dictionary(Of String, Object) From
    {
        {"OrderName", Me.ValueDict("OrderName")},
        {"PickupTime", Me.ValueDict("PickupTime")},
        {"DineIn", Me.ValueDict("DineIn")},
        {"Table", Me.ValueDict("Table")},
        {"Phone", Me.ValueDict("Phone")},
        {"ServerNum", Me.ValueDict("ServerNum")},
        {"ServerName", Me.ValueDict("ServerName")},
        {"CheckDate", Me.ValueDict("CheckDate")},
        {"CheckOpen", Me.ValueDict("CheckOpen")}
    }
        Return dict
    End Function

    Public Function CreateNewOrder(OrderType As Object, check As String) As aclsOrder
        Dim iOrder As New aclsOrder
        iOrder.OrderType = OrderType
        InitializeOrder(check)
        iOrder.check = check
        iOrder.ImportCheckDetails(iOrder, check)
        CreateNewOrder = iOrder
    End Function

    Public Sub InitializeOrder(check As String)
        InitializeCheck(check)
        Me.OrderType.Initialize(check)
    End Sub

    Public Sub TransferCheck(check As String, TransferEmployee As zclsEmployee)
        ReplaceDictValue(Me.ValueDict, "ServerNum", TransferEmployee.ServerNum)
        ReplaceDictValue(Me.ValueDict, "ServerName", TransferEmployee.FirstName)
        UpdateDailyCheckIndex(check, Me.ValueDict)
        Me.OrderType.TransferCheck(check, TransferEmployee)
    End Sub

    Private Function GetTableVal() As Boolean
        ' TODO -- RENAME GetTableVal to GetServiceStyle
        Dim iTable As New zclsTable
        If Match(iTable.Wrap(GetNewMatchObj("CheckNumber", Me.ValueDict("CheckNumber")))) = True Then
            Return True
        End If
        Return False
    End Function

End Class
