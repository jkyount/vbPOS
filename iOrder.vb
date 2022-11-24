Imports Scripting

Module iOrder
    Public Sub SyncValues(Optional Order As aclsOrder = Nothing)
        If Order Is Nothing Then Order = iState.ThisOrder
        Order.ImportCheckDetails(Order, Order.Check)
        Order.OrderType = Order.SameOrderType
    End Sub

    Public Function NewOrderObject(OrderType As Object, check As String) As aclsOrder
        Dim iOrder As New aclsOrder
        NewOrderObject = iOrder.NewOrderObject(OrderType, check)
        iOrder = Nothing
    End Function

    Public Sub SetOrderInfo(Order As aclsOrder, Optional dict As Dictionary(Of String, Object) = Nothing)
        Dim iOrder As New aclsOrder
        iOrder.SetOrderInfo(Order, dict)
        iOrder = Nothing
    End Sub
End Module
