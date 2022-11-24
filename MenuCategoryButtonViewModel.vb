Imports System.Collections.ObjectModel
Imports POS1.Commands
Public Class MenuCategoryButtonViewModel
    Inherits vmBaseButton

    Private pMenuStyle As aclsMenuStyle

    Public Property MenuStyle As aclsMenuStyle
        Set(value As aclsMenuStyle)
            pMenuStyle = value
        End Set
        Get
            Return pMenuStyle
        End Get
    End Property

    Public Overrides Sub ButtonClickCommandEvent()
        'iActiveMenuMessenger.MenuConstructor.MenuStyle = Me.MenuStyle
        iActiveMenuMessenger.ActiveMenuCategory = Me.ID
        iActiveMenuMessenger.RenderTransform = GetLocationSnapshot()
        iActiveMenuMessenger.ActiveMenuName = Me.DisplayName

    End Sub

    Public Sub New(ID As Integer, DisplayName As String)
        MyBase.New(ID, DisplayName)
        Me.MenuStyle = New aclsMenuStyle(GetMenuStyle(New zclsFamily(ID)))
    End Sub

    Private Function GetLocationSnapshot() As TranslateTransform
        Dim tt As New TranslateTransform
        Dim MyPos As Point = Me.GUIInstance.PointToScreen(New Point(0, 0))
        If Me.MenuStyle.StyleDict("AlignToInvoker") Then
            tt.X = MyPos.X
        End If
        Return tt
    End Function

    Public Overrides Function GetNew(ID As Integer, DisplayName As String) As Object
        Throw New NotImplementedException()
    End Function
End Class


