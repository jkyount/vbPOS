Module Globals

    Public Enum MenuStyleEnum
        None = 0
        Condensed = 1
        Expanded = 2
        Component = 3
        Mods = 7
        Discounts = 8
    End Enum





    Public ThisAssembly As String = "POS1"
    Public ThisEmployee As zclsEmployee
    Public ThisTable As New zclsTable
    Public WinColl As New Collection
    Public ThisOrder As New aclsOrder
    Public CurrentCheck As String
    Public iCurrentView As CurrentView
    'Public CurrentSeat As Integer
    Public iState As New State
    Public iActiveMenuMessenger As ActiveMenuMessenger
    Public ThisItem As aclsItem
    Public CItem As New Collection
    Public CurrentParent As Integer
    Public collCheckData As New Collection

    Public LastEntry As New Collection


    Public Function GetNewActiveMenuMessenger() As ActiveMenuMessenger
        iActiveMenuMessenger = New ActiveMenuMessenger
        Return iActiveMenuMessenger
    End Function

    Public Function GetThisEmployee() As zclsEmployee
        Return ThisEmployee
    End Function

    Public Function GetThisTable() As zclsTable
        Return ThisTable
    End Function

    Public Function GetWindows() As Collection
        Return WinColl
    End Function

    Public Function GetThisOrder(iState As State) As aclsOrder

        ThisOrder.ValueDict = GetCheckAttributes(iState.CurrentCheck)
        Return ThisOrder
    End Function

    Public Function GetCurrentCheck() As String
        Return iState.CurrentCheck
    End Function

    Public Function GetCurrentView() As CurrentView
        Return iCurrentView
    End Function

    Public Sub SetInitialState()

        iState = New State

    End Sub
End Module
