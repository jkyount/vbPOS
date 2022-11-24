Imports Microsoft.Vbe.Interop.Forms
Imports Scripting

Module Functions_Order
    Dim Qty As Integer
    Dim currentseat As Integer
    Dim TestColl As Collection
    Dim ClickFlag As Boolean

    Public Sub SetCurrentParentToNothing()
        CurrentParent = 0
    End Sub
    Public Function GetCurrentParent() As Integer
        GetCurrentParent = CurrentParent
    End Function

    Public Function GetCurrentSeat() As Integer
        If iState.CurrentSeat = 0 Then GetCurrentSeat = 1
        GetCurrentSeat = iState.CurrentSeat
    End Function
    Public Sub SetCurrentSeat(val As Integer)
        iState.CurrentSeat = val
    End Sub

    Public Sub ActivateOrderScreen()

        ResetOrderState()
        iState.CurrentSeat = 1
        'CaptureState
    End Sub

    Public Sub Ssfub(ItemID As Integer)
        If iState.ThisEmployee.IDNumber = 0 Then
            'RestoreState
            MsgBox("Session timed out.  Log in to continue.")
            GoToLoginScreen()
            Exit Sub
        End If


        ''''''''Dim bname As String
        ''''''''bname = Application.caller
        ''''''''Dim ItemID As Integer
        ''''''''ItemID = CInt(bname)
        ItemDirector(ItemID)
    End Sub

    Public Sub SetCurrentParent(ID As Integer)
        CurrentParent = ID
        'TODO Implement as below
        'ShowParentFrame(GetItemByID(ID).ItemName)
    End Sub
    Public Sub ItemDirector(ItemID As Integer)

        If ClickFlag = True Then
            Exit Sub
        End If
        ClickFlag = True

        ''''If ItemID = 0 Then ItemID = CInt(Application.caller)
        ThisItem = New aclsItem
        ThisItem.Initialize(ItemID)

        '/////ItemID becomes INTEGER here////'

        Order()
        ThisItem = Nothing
        ClickFlag = False
    End Sub

    Public Sub Order()
        'TODO Disable Component Navigation Buttons As Below
        'ShowShape "ComponentBLOCK"

        ThisItem.OrderRank = ThisItem.AssignOrderRank
        ThisItem.CollID = GetNextCollID(OrderByCollID(DuplicateCItem(CItem)))
        ThisItem.RequiredComponents = ThisItem.GetRequiredComponents

        ThisItem.OrderRank.SetParent
        ThisItem.Children = New bclsChildren
        ThisItem.ItemType.SpclConfig(ThisItem)
        ThisItem.OrderRank.BuildItemCollection

        ThisItem.ItemType.zclsItem_RefreshPreviewWindow
        ThisItem.ItemType.zclsItem_UpdateGUI
        ThisItem.OrderRank.CheckForRequiredComponents(ThisItem)
    End Sub

    Public Function PopItem(ItemID As Integer, item As Object) As aclsItem
        Dim iMenu As New zclsMenu
        Dim ItemDict As New Dictionary(Of String, Object)
        ItemDict = GetValueDict(iMenu.Wrap(GetNewMatchObj("ID", ItemID)))(1)
        item.ItemID = ItemID
        item.ItemName = ItemDict("ItemName")
        item.Price = ItemDict("Price")
        item.Family = ItemDict("Family")
        If Not IsDBNull(ItemDict("Req1")) Then
            item.Req1 = ItemDict("Req1")
            item.Req2 = ItemDict("Req2")
        End If
        If item.Req1 = "" Then
            item.Req1 = item.Req2
            item.Req2 = ""
        End If
        item.IsPrimaryItem = ItemDict("IsPrimaryItem")
        item.Category = ItemDict("Category")
        item.CoolerFlag = ItemDict("Flag")
        item.AlwaysTax = ItemDict("AlwaysTax")
        item.AltScePrice = ItemDict("AltScePrice")
        item.AltPrice = ItemDict("AltPrice")
        item.DiscountPrice = ItemDict("DiscountPrice")
        item.PrintKitchen = ItemDict("PrintKitchen")
        item.PrintPantry = ItemDict("PrintPantry")
        item.ItemOptions = ItemDict("ItemOptions")

        item.seat = GetCurrentSeat()
        Return item
    End Function

    'Public Sub GetItemCell()
    '    If Sheet1.range("CheckOrientationCell").value = "" Then
    '        ItemCell = Sheet1.range("CheckOrientationCell")
    '        Exit Sub
    '    End If
    '    ItemCell = Sheet1.range("CheckOrientationCell").Offset(100, 0).End(xlUp).Offset(1, 0)
    'End Sub

    Public Sub AdvCell()
        'ItemCell = ItemCell.Offset(1, 0)
    End Sub
    'Public Sub AdvRange(rg As range)
    '    rg = rg.Offset(1, 0)
    'End Sub

    'Public Sub SetEndOfCheck()
    '    EndOfCheck = ItemCell.row
    'End Sub


    'Public Sub WholeTopping()
    '    CItem("1").ItemType.ToppingArea = 0
    '    Sheet1.Shapes("PzaWhole").line.weight = 3
    '    Sheet1.Shapes("PzaHalf").line.weight = 1
    'End Sub
    'Public Sub HalfTopping()
    '    CItem("1").ItemType.ToppingArea = 1
    '    Sheet1.Shapes("PzaWhole").line.weight = 1
    '    Sheet1.Shapes("PzaHalf").line.weight = 3
    'End Sub
    Public Sub SetQuantity()
        Dim QtyValidate As Object
        QtyValidate = InputBox("Set quantity for next item Ordered:", , "1")
        If IsNumeric(QtyValidate) Then
            If CInt(QtyValidate) > 0 Then
                Qty = QtyValidate
                Exit Sub
            End If
        End If
        MsgBox("Please enter a positive numeric value.")

    End Sub

    Public Sub SetToDineIn(check As String)
        Dim iIndex As New zclsDailyCheckIndex
        Update(iIndex.Wrap(GetNewUpdateObj(, check, "DineIn", True)))
        'TODO Implement As Below
        'SetOrderTypeIndicator("DineIn")
        UpdateTax(iState.CurrentCheck)
        iIndex = Nothing
    End Sub
    Public Sub SetToCarryout(check As String)
        Dim iIndex As New zclsDailyCheckIndex
        Update(iIndex.Wrap(GetNewUpdateObj(, check, "DineIn", False)))
        'TODO Implement As Below
        'SetOrderTypeIndicator("Carryout")
        UpdateTax(iState.CurrentCheck)
        iIndex = Nothing
    End Sub

    Public Sub CheckQueueForRequiredComponents(coll As Collection)
        Dim i As Integer
        For i = coll.Count To 1 Step -1
            Dim Reqs As Object
            Reqs = AllRequirementsFullfilled(coll(i))
            If Not TypeOf Reqs Is Boolean Then

                'If Not Reqs = True Then
                MsgBox("Required component " & Reqs & " not selected.")
                    SetCurrentParent(coll(i).CollID)
                    Exit Sub
                'End If
            End If
        Next i

        Dim FormattedColl As Collection
        FormattedColl = FormatItemCollection(coll)
        OrderQueuedItems(FormattedColl)
        ResetOrderState()
    End Sub

    Public Function AllRequirementsFullfilled(item As aclsItem) As Object
        If Not item.RequiredComponents.Count = 0 Then
            For Each key As Object In item.RequiredComponents.Keys
                Dim RequiredComponent As String
                RequiredComponent = item.RequiredComponents(key)
                If item.ValidateRequirements(item, RequiredComponent) = False Then
                    'TODO Implement As Below
                    iActiveMenuMessenger.ActiveMenuCategory = New zclsFamily(RequiredComponent).FamilyID
                    'ShowComponentFrame(RequiredComponent)
                    'DisplayQuickMods(RequiredComponent)
                    'ShowParentFrame(item.ItemName)
                    AllRequirementsFullfilled = RequiredComponent
                    Exit Function
                End If
            Next key
        End If
        AllRequirementsFullfilled = True
    End Function



    'TODO Function signature should be base class for item
    Public Function RequiresComponents(item As Object) As Boolean
        If item.Req1 = "" Then
            RequiresComponents = False
            Exit Function
        End If
        RequiresComponents = True
    End Function

    Public Function FormatSides(coll As Collection) As Collection
        For Each member As aclsItem In coll
            If Not member.Children.coll.Count = 0 Then
                For Each child As bclsChild In member.Children.coll
                    Dim ChildItem As aclsItem
                    ChildItem = child.item
                    If ChildItem.Family = "Drsng" Or ChildItem.Family = "Sce" Then
                        member.ItemName = member.ItemName & "  /  " & LTrim(ChildItem.ItemName)
                        member.Price = (member.Price) + (ChildItem.Price)
                        coll.Remove(CStr(child.ID))
                    End If
                Next child
            End If
        Next member
        FormatSides = coll
    End Function

    Public Sub CLICK_Sheet1_btnDONE()
        CheckQueueForRequiredComponents(CItem)
    End Sub
    Public Sub CancelOrder()
        ''''RestoreState
        Dim CheckDetail As New zclsDailyCheckDetail
        DeleteMatch(CheckDetail.Wrap(GetNewMatchObj("CheckNumber", iState.CurrentCheck, "Sent", False)))
        If Match(CheckDetail.Wrap(GetNewMatchObj("CheckNumber", iState.CurrentCheck, "Sent", True))) = False Then
            ThisOrder.OrderType.CancelOrder

            SetCheckUnused(iState.CurrentCheck)
            Dim CheckIndex As New zclsDailyCheckIndex
            DeleteMatch(CheckIndex.Wrap(GetNewMatchObj("CheckNumber", iState.CurrentCheck)))
        End If
        'TODO Implement as below
        'ActivateHomeScreen
    End Sub
    Public Sub OrderSend(check As String)
        SendCurrentItems(check)
        DailyCheckIndex_CalculateTotals(check)
    End Sub



    Public Sub OrderQueuedItems(coll As Collection)
        Dim q As Integer
        q = 1
        GetQty()
        Dim DailyCheckDetail As New zclsDailyCheckDetail
        Do While q <= Qty
            DailyCheckDetail.AddCurrentItems(coll)
            q = q + 1
        Loop
        collCheckData = RecallCheckLines(iState.CurrentCheck)
        'TODO Implement as Below
        'WriteCheckLines Sheet1.range("CheckRange"), collCheckData
        'PopGuiWithCheckAttributes GetCheckAttributes(CurrentCheck), Sheet1
        SetLastEntry(coll)
        ResetOrderState()

        DailyCheckDetail = Nothing


    End Sub
    Public Sub ResetOrderState()
        iActiveMenuMessenger.Reset()
        iState.ThisOrder = New aclsOrder(iState.CurrentCheck)
        SyncValues()
        iCurrentView.ViewModel.SetContent(New pgOrder_TopBar, New pgOrder_Content)
        iState.CurrentSeat = 1
        Qty = 1
    End Sub

    Public Sub GetQty()
        If IsDBNull(Qty) Or Qty = 0 Then Qty = 1
    End Sub

    Public Sub SetLastEntry(coll As Collection)
        If coll.Count = 0 Then Exit Sub
        LastEntry = DuplicateCollection(coll)
        ClearCollection(CItem)
    End Sub

    Public Sub RepeatLastEntry()
        'MsgBox "Unavailable"
        'Exit Sub
        OrderQueuedItems(LastEntry)
    End Sub

    Public Sub CloseCheck()
        ''''RestoreState
        OrderSend(iState.CurrentCheck)
        DailyCheckIndex_CloseOrder(iState.CurrentCheck)
        ThisOrder.OrderType.CloseCheck(iState.CurrentCheck)
        MsgBox("Check " & iState.CurrentCheck & " closed.")
        'TODO Implement as below
        'ActivateHomeScreen
    End Sub

    'TODO Implement as below
    'Public Sub QuickMod()
    '    Dim bname As String
    '    bname = Application.caller
    '    Sheet1.SpecialInstructionText.value = Sheet1.Shapes(bname).TextFrame.Characters.text
    '    Dim ItemID As Integer
    '    ItemID = GetNextItemID("SpclInstruction")
    '    ItemDirector(ItemID)
    '    Sheet1.SpecialInstructionText.value = ""
    'End Sub

    Public Function MissingComponents() As Boolean
        Dim i As Integer
        For i = CItem.Count To 1 Step -1
            Dim Reqs As Object
            Reqs = AllRequirementsFullfilled(CItem(i))
            'If Not CType(Reqs, Boolean) = True Then
            If Not TypeOf Reqs Is Boolean Then
                    SetCurrentParent(CItem(i).CollID)
                    MissingComponents = True
                    Exit Function
                End If
        Next i
        MissingComponents = False
    End Function

    Public Sub ScrollParents()

        Dim CurrentParentCollIndex As Integer
        Dim i As Integer
        For i = 1 To CItem.Count
            If CItem(i).CollID = CurrentParent Then CurrentParentCollIndex = i
        Next i

        If Not CurrentParentCollIndex = CItem.Count Then
            SetCurrentParent(CItem(CurrentParentCollIndex + 1).CollID)
            Exit Sub
        End If

        SetCurrentParent(CItem(1).CollID)



    End Sub
End Module
