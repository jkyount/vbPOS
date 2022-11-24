Imports Scripting

Module iDailyCheckIndex

    Public Sub DailyCheckIndex_CalculateTotals(check As String) '<------- REVISED VERSION??
        Dim Ord As New aclsOrder
        Dim iDetail As New zclsDailyCheckDetail
        Dim dict As New Dictionary(Of String, Object)
        Dim iIndex As New zclsDailyCheckIndex
        Ord.ImportCheckDetails(Ord, check)
        Ord.OrderType = Ord.SameOrderType

        dict.Add("SubTotal", SumMatch(iDetail.Wrap(GetNewMatchObj(, check)), "Price"))

        dict.Add("Beer", SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Category", "Beer")), "Price"))
        dict.Add("Wine", SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Category", "Wine")), "Price"))
        dict.Add("Discount", SumMatch(iDetail.Wrap(GetNewMatchObj(, check, "Family", "Discounts")), "Price"))
        dict.Add("FoodIn", Ord.OrderType.GetFoodInTotal(check))
        dict.Add("Carryout", Ord.OrderType.GetCarryoutTotal(check))
        UpdateFromDict(iIndex.Wrap(GetNewMatchObj(, check)), dict)
        UpdateTax(check)
        UpdateTotal(check)
    End Sub

    Public Sub UpdateTax(check As String)
        Dim iDailyCheckIndex As New zclsDailyCheckIndex
        iDailyCheckIndex.UpdateTax(check)
        iDailyCheckIndex = Nothing
    End Sub

    Public Sub UpdateTotal(check As String)
        Dim iDailyCheckIndex As New zclsDailyCheckIndex
        iDailyCheckIndex.UpdateTotal(check)
        iDailyCheckIndex = Nothing
    End Sub







    Public Sub DailyCheckIndex_CloseOrder(check As String)
        Dim iDailyCheckIndex As New zclsDailyCheckIndex
        iDailyCheckIndex.CloseOrder(check)
        iDailyCheckIndex = Nothing
    End Sub

    Public Function GetChecks(Optional Where As String = "", Optional Equals As Object = Nothing, Optional AndWhere As String = "", Optional Equals2 As Object = Nothing) As Array
        Dim iDataObject As New aclsDataObject
        Dim iDailyCheckIndex As New zclsDailyCheckIndex
        iDataObject = iDailyCheckIndex.Wrap(GetNewMatchObj(Where, Equals, AndWhere, Equals2))
        GetChecks = iDailyCheckIndex.FormatRecordset(GetRecordsetMatch(iDataObject, ConstructMatchQuery(iDataObject)))
    End Function

    'Public Function GetCheckAttributes(check As String) As Scripting.Dictionary
    '    Dim iIndex As New zclsDailyCheckIndex
    '    GetCheckAttributes = GetValueDict(iIndex.Wrap(GetNewMatchObj("CheckNumber", check)))(1)
    '    iIndex = Nothing
    'End Function

    Public Function GetCheckAttributes(check As String) As Dictionary(Of String, Object)
        Dim iIndex As New zclsDailyCheckIndex
        GetCheckAttributes = GetValueDict(iIndex.Wrap(GetNewMatchObj("CheckNumber", check)))(1)
        iIndex = Nothing
    End Function

    Public Sub UpdateDailyCheckIndex(check As String, Params As Dictionary(Of String, Object))
        Dim iDailyCheckIndex As New zclsDailyCheckIndex
        iDailyCheckIndex.UpdateDailyCheckIndex(check, Params)
        iDailyCheckIndex = Nothing
    End Sub
End Module
