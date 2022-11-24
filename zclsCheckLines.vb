Imports System.Collections.ObjectModel
Imports Microsoft.Vbe.Interop.Forms
Imports Scripting

Public Class zclsCheckLines
    'Inherits zclsBase
    Private pRow As Integer
    Private pPrimary As Boolean
    Private pData() As Object
    Private pPrintKitchen As Boolean
    Private pPrintPantry As Boolean
    Private pSeat As Integer
    Private pGuiRow As Integer
    Private pCashPayment As Double
    Private pChargePayment As Double
    Private pGiftCertPayment As Double
    Private pLocalGroup As Integer
    Private pCoolerFlag As Boolean
    Private pItemID As Integer
    Private pEntityGroup As Integer



    Public dbCheck As New ADODB.Connection
    Public rsCheck As New ADODB.Recordset

    Public ReadOnly Property Db As String
        Get
            Return GetDb()
        End Get
    End Property

    Public Property Row As Integer
        Set(value As Integer)
            pRow = value
        End Set
        Get
            Return pRow
        End Get
    End Property

    Public Property GuiRow As Integer
        Set(value As Integer)
            pGuiRow = value
        End Set
        Get
            Return pGuiRow
        End Get
    End Property

    Public Property Seat As Integer
        Set(value As Integer)
            pSeat = value
        End Set
        Get
            Return pSeat
        End Get
    End Property

    Public Property Primary As Boolean
        Set(value As Boolean)
            pPrimary = value
        End Set
        Get
            Return pPrimary
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

    Public Property Data As Object
        Set(value As Object)
            pData = value
            'OnPropertyChanged(NameOf(Data))
        End Set
        Get
            Return pData
        End Get
    End Property

    Public Property LocalGroup As Integer
        Set(value As Integer)
            pLocalGroup = value
        End Set
        Get
            Return pLocalGroup
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

    Public Property EntityGroup As Integer
        Set(value As Integer)
            pEntityGroup = value
        End Set
        Get
            Return pEntityGroup
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







    Public Function Wrap(obj As aclsDataObject) As aclsDataObject
        Dim iDataObj As New aclsDataObject
        iDataObj = obj
        iDataObj.Rs = Me.GetRs
        iDataObj.Conn = Me.GetConn
        iDataObj.Db = Me.GetDb
        iDataObj.DbFile = Me.GetDbFile
        Wrap = iDataObj
        iDataObj = Nothing
    End Function

    Public Function GetDb() As String
        GetDb = "DailyCheckDetail"
    End Function
    Public Function GetDbFile() As String
        GetDbFile = "CheckDb"
    End Function

    Public Function GetConn() As ADODB.Connection
        GetConn = dbCheck
    End Function

    Public Function GetRs() As ADODB.Recordset
        GetRs = rsCheck
    End Function

    Public Sub OpenDbs()
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(iDataObject)
        iDataObject.OpenDbs(iDataObject)
        iDataObject = Nothing
    End Sub

    Public Sub CloseDbs()
        Dim iDataObject As New aclsDataObject
        iDataObject = Wrap(iDataObject)
        iDataObject.CloseDbs(iDataObject)
        iDataObject = Nothing
    End Sub



    '==========================================================================

    Public Sub New()

    End Sub

    Public Sub New(data As Array)
        Me.Data = data
    End Sub
    Public Function DefineWriteLines(coll As Collection, WriteLines As Object) As Collection
        DefineWriteLines = pDefineWriteLines(coll, WriteLines)
    End Function
    Public Function pDefineWriteLines(coll As Collection, WriteLines As Object) As Collection

        Dim i As Integer
        Dim arr() As Object
        Dim arr2() As Object
        Dim member As Object
        For Each member In coll
            arr2 = member.Data
            ReDim arr(0 To UBound(WriteLines))
            For i = 0 To UBound(WriteLines)
                arr(i) = arr2(WriteLines(i))
            Next i
            member.Data = arr
        Next member

        pDefineWriteLines = coll
    End Function

    Public Function CreateNew(row As Integer, ValueDict As Dictionary(Of String, Object)) As zclsCheckLines
        Dim x As New zclsCheckLines
        Dim TempArray(0 To 4) As Object
        x.Row = row
        x.ItemID = ValueDict("ItemID")
        TempArray(0) = "PlaceHolder"
        TempArray(1) = ValueDict("ItemIndicator")
        TempArray(2) = ValueDict("ItemName")
        TempArray(3) = FormatCurrency(ValueDict("Price"), 2)
        TempArray(4) = ValueDict("ItemID")
        x.PrintKitchen = ValueDict("PrintKitchen")
        x.PrintPantry = ValueDict("PrintPantry")
        x.Seat = ValueDict("Seat")
        x.LocalGroup = ValueDict("LocalGroup")
        x.Primary = ValueDict("IsPrimaryItem")
        x.EntityGroup = ValueDict("EntityGroup")
        x.Data = TempArray
        CreateNew = x
    End Function
    Public Function GetCheckLines(check As String, ValueDict As Collection) As Collection
        GetCheckLines = pGetCheckLines(check, ValueDict)
    End Function
    Public Function pGetCheckLines(check As String, ValueDict As Collection) As Collection
        'If Not IsEmpty(CheckArray) Then

        Dim line As New zclsCheckLines
        Dim coll As New Collection
        If Not ValueDict(1).Count = 0 Then
            Dim i As Integer
            For i = 1 To ValueDict.Count
                line = line.CreateNew(coll.Count + 1, ValueDict(i))
                coll.Add(line, ("Line" & i))
            Next i
        End If
        pGetCheckLines = coll
        coll = Nothing
        ValueDict = Nothing
    End Function

    Public Sub WriteCheckLines(range As Object, coll As Collection)
        Dim i As Integer, SeatLines As Integer, GuiRow As Integer, seat As Integer
        Dim seatarray(0 To 2) As String
        If coll.Count = 0 Then
            range.value = ""
            'MsgBox "Attempted to write an empty collection."
            Exit Sub
        End If
        range.value = ""
        SeatLines = 1
        seat = coll("Line1").seat
        GuiRow = 2
        seatarray(0) = "- - - - - - -"
        seatarray(1) = "- - - - - - - Seat " & seat & "- - - - - - - - -"
        seatarray(2) = "- - - - - - -"
        range.Rows(1).value = seatarray
        For i = 1 To coll.Count
            If Not coll("Line" & i).seat = seat Then
                seat = coll("Line" & i).seat
                seatarray(1) = "- - - - - - - Seat " & seat & "- - - - - - - - -"
                range.Rows(i + SeatLines).value = seatarray
                SeatLines = SeatLines + 1
            End If
            range.Rows(i + SeatLines).value = coll("Line" & i).Data()
        Next i

    End Sub
End Class
