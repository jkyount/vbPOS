Imports Scripting

Public Class zclsDailyCheckIndex
    Inherits DBObject
    Public dbDailyCheckIndex As New ADODB.Connection
    Public rsDailyCheckIndex As New ADODB.Recordset



    '==========================================================================
    '==========================================================================





    Public Overrides Function GetDb() As String
        GetDb = "DailyCheckIndex"
    End Function

    Public Overrides Function GetDbFile() As String
        GetDbFile = "CheckDb"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbDailyCheckIndex
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsDailyCheckIndex
    End Function

    Public Function GetArchive() As String
        GetArchive = "ArchivedCheckIndex"
    End Function
    Public Function GetArchiveDbFile() As String
        GetArchiveDbFile = "ReportsDb"
    End Function



    '==========================================================================






    '==========================================================================




    'UNCOMMENT THESE AFTER IMPORTING aclsOrder CLASS ------

    Public Sub UpdateTax(check As String)
        pUpdateTax(check)
    End Sub

    Private Sub pUpdateTax(check As String)
        Dim Ord As New aclsOrder
        Ord.ImportCheckDetails(Ord, check)
        Ord.OrderType = Ord.SameOrderType

        Update(Wrap(GetNewUpdateObj(, check, "Tax", Ord.OrderType.GetTax(check))))
        Ord = Nothing

    End Sub


    Public Sub UpdateTotal(check As String)
        pUpdateTotal(check)
    End Sub

    Private Sub pUpdateTotal(check As String)
        OpenDbs()
        rsDailyCheckIndex.Source = "SELECT * From DailyCheckIndex WHERE CheckNumber = """ & check & """"
        rsDailyCheckIndex.Open()
        rsDailyCheckIndex.Fields("Total").Value = rsDailyCheckIndex.Fields("Subtotal").Value + rsDailyCheckIndex.Fields("Tax").Value + rsDailyCheckIndex.Fields("ServiceCharge").Value
        rsDailyCheckIndex.Update()
        rsDailyCheckIndex.Close()
        dbDailyCheckIndex.Close()
    End Sub


    Public Sub UpdateDailyCheckIndex(check As String, Params As Dictionary(Of String, Object))
        pUpdateDailyCheckIndex(check, Params)
    End Sub
    Private Sub pUpdateDailyCheckIndex(check As String, Params As Dictionary(Of String, Object))
        OpenDbs()
        rsDailyCheckIndex.Source = "SELECT * From DailyCheckIndex WHERE CheckNumber = """ & check & """"
        rsDailyCheckIndex.Open()

        For Each member As Object In Params
            If Not member = "Taxable" Then
                rsDailyCheckIndex.Fields(member).Value = Params(member)
            End If
        Next member
        rsDailyCheckIndex.Fields("CheckNumber").Value = check

        rsDailyCheckIndex.Update()
        rsDailyCheckIndex.Close()
        dbDailyCheckIndex.Close()
    End Sub

    Public Function FormatRecordset(ChecksRecordset As ADODB.Recordset) As Array

        FormatRecordset = pFormatRecordset(ChecksRecordset)
    End Function
    Private Function pFormatRecordset(ChecksRecordset As ADODB.Recordset) As Array
        'OpenDbs
        'rsDailyCheckIndex.Source = strQuery
        'ChecksRecordset.Open
        Dim arrCheck As Object
        Dim arrTotal As Object
        Dim arrName As Object
        Dim arrTime As Object
        Dim arrServer As Object
        Dim i As Integer


        If ChecksRecordset.RecordCount = 0 Then

            Dim EmptyArray(0 To 0, 0 To 4)
            EmptyArray(0, 0) = ""
            EmptyArray(0, 1) = "No"
            EmptyArray(0, 2) = "Open"
            EmptyArray(0, 3) = "Checks"
            EmptyArray(0, 4) = ""
            pFormatRecordset = EmptyArray
            CloseDbs()
            Exit Function
        End If

        arrCheck = ChecksRecordset.GetRows(, , "CheckNumber")
        rsDailyCheckIndex.MoveFirst()
        arrName = ChecksRecordset.GetRows(, , "OrderName")
        rsDailyCheckIndex.MoveFirst()
        arrTime = ChecksRecordset.GetRows(, , "PickupTime")
        rsDailyCheckIndex.MoveFirst()
        arrTotal = ChecksRecordset.GetRows(, , "SubTotal")
        rsDailyCheckIndex.MoveFirst()
        arrServer = ChecksRecordset.GetRows(, , "ServerName")
        CloseDbs()

        Dim TempArray(0 To UBound(arrCheck, 2), 0 To 4)
        For i = 0 To UBound(arrCheck, 2)
            TempArray(i, 0) = arrCheck(0, i)
        Next i
        For i = 0 To UBound(arrCheck, 2)
            TempArray(i, 1) = arrName(0, i)
        Next i
        For i = 0 To UBound(arrCheck, 2)
            TempArray(i, 2) = arrTime(0, i)
        Next i
        For i = 0 To UBound(arrCheck, 2)
            TempArray(i, 3) = arrTotal(0, i)
        Next i
        For i = 0 To UBound(arrCheck, 2)
            TempArray(i, 4) = arrServer(0, i)
        Next i
        pFormatRecordset = TempArray
    End Function

    Public Sub CloseOrder(check As String)
        pCloseOrder(check)
    End Sub
    Private Sub pCloseOrder(check As String)
        Dim iIndex As New zclsDailyCheckIndex

        Update(iIndex.Wrap(GetNewUpdateObj(, check, "Closed", True)))
        Update(iIndex.Wrap(GetNewUpdateObj(, check, "CheckClose", Format(Now, "hh:mm"))))

    End Sub







End Class
