Imports Scripting
Imports System.Runtime.Remoting

Public Class aclsMenuStyle
    Inherits DBObject
    Public dbMenuStyle As New ADODB.Connection
    Public rsMenuStyle As New ADODB.Recordset

    Private pStyleDict As Dictionary(Of String, Object)
    Private pMenuComponents As aclsMenuComponents

    Public Property MenuComponents As aclsMenuComponents
        Set(value As aclsMenuComponents)
            pMenuComponents = value
        End Set
        Get
            Return pMenuComponents
        End Get
    End Property

    Public Property StyleDict As Dictionary(Of String, Object)
        Set(value As Dictionary(Of String, Object))
            pStyleDict = value
        End Set
        Get
            Return pStyleDict
        End Get
    End Property

#Region "Overrides"


    Public Overrides Function GetDb() As String
        GetDb = "MenuStyle"
    End Function

    Public Overrides Function GetDbFile() As String
        GetDbFile = "Menu"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbMenuStyle
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsMenuStyle
    End Function

#End Region

    Public Sub New()

    End Sub

    Public Sub New(MenuStyle As Integer)
        Me.StyleDict = GetStyleDict(MenuStyle)
        Me.MenuComponents = New aclsMenuComponents(StyleDict("GridType"), StyleDict("ButtonType"))
    End Sub

    Public Function GetNewMenuStyleObj(MenuStyle As String) As aclsMenuStyle
        Return Nothing
        MsgBox("Converted this function to be called in class constructor")
    End Function

    Private Function GetStyleDict(MenuStyle As Integer) As Dictionary(Of String, Object)
        Dim iDataObj As aclsDataObject = Wrap(GetNewMatchObj("ID", MenuStyle))
        Return CDictCollection(GetRecordsetMatch(iDataObj, ConstructMatchQuery(iDataObj)))(1)
    End Function

    Public Function GetMenuStyles() As Object

        Dim iDataObj As New aclsDataObject
        iDataObj = Me.Wrap(GetNewMatchObj)
        Dim qry As String
        qry = "SELECT MenuStyle FROM MenuStyle"
        GetMenuStyles = RsToArray(GetRecordsetMatch(iDataObj, qry))
        iDataObj.CloseDbs(iDataObj)
        iDataObj = Nothing
    End Function
End Class


Public Class aclsMenuComponents

    Private pGrid As UserControl
    Private pButtonType As vmBaseButton


    Public Property Grid As UserControl
        Set(value As UserControl)
            pGrid = value
        End Set
        Get
            Return pGrid
        End Get
    End Property
    Public Property ButtonType As vmBaseButton
        Set(value As vmBaseButton)
            pButtonType = value
        End Set
        Get
            Return pButtonType
        End Get
    End Property

    Public Sub New(GridKeyword As String, ButtonKeyword As String)
        Me.Grid = Activator.CreateInstance(vbNullString, ThisAssembly & "." & GridKeyword).Unwrap
        Me.ButtonType = Activator.CreateInstance(vbNullString, ThisAssembly & "." & ButtonKeyword).Unwrap
    End Sub
End Class