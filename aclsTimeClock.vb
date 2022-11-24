Imports Microsoft.Office.Core
Imports Scripting

Public Class aclsTimeClock


    Public dbTimeClock As New ADODB.Connection
    Public rsTimeClock As New ADODB.Recordset

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
        GetDb = "TimeClock"
    End Function

    Public Function GetDbFile() As String
        GetDbFile = "Employee"
    End Function

    Public Function GetConn() As ADODB.Connection
        GetConn = dbTimeClock
    End Function

    Public Function GetRs() As ADODB.Recordset
        GetRs = rsTimeClock
    End Function

    Public Function GetArchive() As String
        GetArchive = "TimeClock"
    End Function
    Public Function GetArchiveDbFile() As String
        GetArchiveDbFile = "Employee"
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



    Public Sub ClockIn(ID As Long)
        Dim k As DateTime
        k = Format(Now, "MM/dd/yy HH:mm")
        AddNewRecord(Me, GetTimeClockDict(ID, k))
        MsgBox("Clocked in at" & k)

    End Sub

    Private Function GetTimeClockDict(EmployeeID As Long, k As DateTime) As Dictionary(Of String, Object)
        Dim dict As New Dictionary(Of String, Object)
        Dim Employee As New zclsEmployee(EmployeeID)
        With dict
            .Add("IdNumber", EmployeeID)
            .Add("ClockDate", Format(Now, "MMddyy"))
            .Add("FirstName", Employee.FirstName)
            .Add("LastName", Employee.LastName)
            .Add("ServerNum", Employee.ServerNum)
            .Add("In", Format(k, "MM/dd/yy HH:mm"))
        End With
        GetTimeClockDict = dict
        dict = Nothing

    End Function

    Public Sub ClockOut(EmployeeID As Long)
        Dim ExtendQuery As String = "Out = False AND IDNumber"
        Dim OutTime As Date = Format(Now, "MM/dd/yy hh:mm")
        Dim OutDate As Date = Format(Date.Today, "MM/dd/yy")
        Dim RecordIndex As Integer = ValueMatch(Wrap(GetNewMatchObj(ExtendQuery, EmployeeID)), "ID")

        Dim InDate As DateTime = Format(ValueMatch(Wrap(GetNewMatchObj("ID", RecordIndex)), "In"), "MM/dd/yy")
        Dim EndOfDayClockOut As DateTime = CDate(" 11:59 PM" & InDate)
        Dim StartOfDayClockIn As DateTime = CDate("12:01 AM" & OutDate)

        If Not OutDate = InDate Then
            Update(Wrap(GetNewUpdateObj("ID", RecordIndex, "Out", EndOfDayClockOut)))
            Update(Wrap(GetNewUpdateObj("ID", RecordIndex, "OutDate", Format(InDate, "MMddyy"))))
            AddNewRecord(Me, GetTimeClockDict(EmployeeID, StartOfDayClockIn))
            RecordIndex = ValueMatch(Wrap(GetNewMatchObj(ExtendQuery, EmployeeID)), "ID")
        End If
        Update(Wrap(GetNewUpdateObj("ID", RecordIndex, "Out", OutTime)))
        Update(Wrap(GetNewUpdateObj("ID", RecordIndex, "OutDate", Format(OutDate, "MMddyy"))))




        MsgBox("Clocked out at " & OutTime)



    End Sub

    Public Sub ClockOutAll()

        Dim coll As New Collection
        Dim rs As New ADODB.Recordset

        Dim ExtendQuery As String
        Dim k As DateTime
        k = Format(Now, "hh:mm")
        ExtendQuery = "Out = False AND NOT IDNumber"
        Dim iDataObj As New aclsDataObject
        iDataObj = Wrap(GetNewMatchObj(ExtendQuery, 0))
        rs = GetRecordsetMatch(iDataObj, ConstructMatchQuery(iDataObj))
        ClearCollection(coll)
        Do
            coll.Add(rs.Fields("IDNumber").Value)
            rs.MoveNext()
        Loop Until rs.EOF

        Me.CloseDbs()

        Dim ID As Integer
        For ID = 1 To coll.Count
            Me.ClockOut(coll(ID))
        Next ID
        MsgBox("All employees have been clocked out at " & k & ".")

    End Sub

End Class
