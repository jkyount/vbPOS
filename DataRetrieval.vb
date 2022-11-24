Module DataRetrieval
    Dim rs As ADODB.Recordset


    Public Function GetNewMatchObj(Optional Where As String = "CheckNumber", Optional Equals As Object = Nothing, Optional AndWhere As String = "", Optional Equals2 As Object = Nothing) As aclsDataObject
        Dim iDataObject As New aclsDataObject

        Return iDataObject.GetNewMatchObj(Where, Equals, AndWhere, Equals2)
        iDataObject = Nothing
    End Function

    Public Function GetNewUpdateObj(Optional Where As String = "CheckNumber", Optional Equals As Object = Nothing, Optional UpdateField As String = "", Optional UpdateValue As Object = Nothing) As aclsDataObject
        Dim iDataObject As New aclsDataObject

        Return iDataObject.GetNewUpdateObj(Where, Equals, UpdateField, UpdateValue)
        iDataObject = Nothing
    End Function

    Public Function GetNewArchiveObj(obj As Object, iDataObj As aclsDataObject) As aclsDataObject

        Return iDataObj.GetNewArchiveObj(obj, iDataObj)

    End Function

    Public Function Match(DataObject As aclsDataObject) As Boolean

        rs = DataObject.GetMatch(DataObject, ConstructMatchQuery(DataObject))
        If Not rs.EOF = True Then
            Match = True
        End If
        DataObject.CloseDbs(DataObject)
        rs = Nothing
    End Function

    Public Function GetMatch(DataObject As aclsDataObject) As Object
        '5/21
        If Match(DataObject) = False Then
            MsgBox("Invalid match parameters.  No match found.")
            Exit Function
        End If
        rs = DataObject.GetMatch(DataObject, ConstructMatchQuery(DataObject))
        GetMatch = GetResult(rs)
        'Dim ResultArray() As Object
        'If rs.RecordCount = 0 Then
        '    GetMatch = Empty
        '    DataObject.CloseDbs DataObject
        '    Exit Function
        'End If
        'ReDim ResultArray(1 To rs.RecordCount)
        'Dim i As Integer
        'For i = 1 To rs.RecordCount
        '    ResultArray(i) = rs.GetRows(1)
        'Next i
        'GetMatch = ResultArray
        DataObject.CloseDbs(DataObject)
        rs = Nothing
    End Function

    Public Function GetRecordsetMatch(obj As aclsDataObject, Optional Query As String = "") As ADODB.RecordSet
        If Query = "" Then
            Query = ConstructMatchQuery(obj)
        End If
        Dim iDataObject As New aclsDataObject
        rs = iDataObject.GetMatch(obj, Query)
        GetRecordsetMatch = rs
        rs = Nothing

    End Function

    Public Function CDictCollection(rs As ADODB.RecordSet) As Collection
        Dim coll As New Collection
        Dim dict As Dictionary(Of String, Object)
        Dim name As String
        Dim value As Object
        If rs.RecordCount = 0 Then
            dict = Nothing
            coll.Add(dict)
            Return coll
            Exit Function
        End If

        Dim fld As Object
        rs.MoveFirst
        Do Until rs.EOF
            dict = New Dictionary(Of String, Object)
            For Each fld In rs.Fields
                name = fld.name
                value = fld.value
                dict.Add(name, value)
            Next fld
            coll.Add(dict)
            rs.MoveNext
        Loop
        Return coll

    End Function

    Public Function CountMatch(obj As aclsDataObject) As Integer
        rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
        CountMatch = rs.RecordCount
        obj.CloseDbs(obj)
        rs = Nothing
    End Function

    Public Function RsToArray(RecordSet As ADODB.Recordset) As Array
        If Not RecordSet.RecordCount = 0 Then
            Return RecordSet.GetRows
        End If
    End Function

    Public Sub DeleteMatch(obj As aclsDataObject)
        rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
        If Not rs.RecordCount = 0 Then
            Do Until rs.EOF
                rs.Delete
                rs.MoveNext
            Loop
        End If
        obj.CloseDbs(obj)
        rs = Nothing
    End Sub

    Public Function SumMatch(obj As aclsDataObject, Field As String) As Double
        rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
        Dim Total As Double
        If Not rs.RecordCount = 0 Then
            Do Until rs.EOF
                Total = Total + rs.Fields(Field).Value
                rs.MoveNext
            Loop
            obj.CloseDbs(obj)
            rs = Nothing
            SumMatch = Total
            Exit Function
        End If
        SumMatch = 0
        obj.CloseDbs(obj)
        rs = Nothing
    End Function

    Public Function ValueMatch(obj As aclsDataObject, Field As String) As Object
        rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
        If Not rs.RecordCount = 0 Then
            ValueMatch = rs.Fields(Field).Value
        End If
        obj.CloseDbs(obj)
        rs = Nothing
    End Function

    Public Sub AddToMatch(obj As aclsDataObject)
        Dim i As Integer
        Dim UpdateField As String
        Dim Amount As String
        Amount = obj.Value2
        UpdateField = obj.Field2

        rs = GetRecordsetMatch(obj, ConstructUpdateQuery(obj))
        If Not rs.RecordCount = 0 Then
            For i = 1 To rs.RecordCount
                rs.Fields(UpdateField).value = rs.Fields(UpdateField).value + Amount
                rs.Update
                rs.MoveNext
            Next i
        End If
        obj.CloseDbs(obj)
        rs = Nothing
    End Sub

    Public Function FilteredMatch(obj As aclsDataObject, ParamArray SelectedFields() As Array) As Object
        '5/21
        Dim arr(0 To UBound(SelectedFields)) As Array
        'ReDim arr(0 To UBound(SelectedFields)) As Variant
        arr = SelectedFields
        rs = obj.GetMatch(obj, ConstructFilteredQuery(obj, arr))
        FilteredMatch = GetResult(rs)
        'If rs.RecordCount = 0 Then
        '    FilteredMatch = False
        '    obj.CloseDbs obj
        '     rs = Nothing
        '    Exit Function
        'End If
        'Dim ResultArray() As Variant
        'ReDim ResultArray(1 To rs.RecordCount)
        'Dim i As Integer
        'For i = 1 To rs.RecordCount
        '    ResultArray(i) = rs.GetRows(1)
        'Next i
        'FilteredMatch = ResultArray
        obj.CloseDbs(obj)
        rs = Nothing
    End Function

    Public Function FltrdOrdrdMtch(obj As aclsDataObject, OrderBy As String, ParamArray SelectedFields() As Array) As Object
        '5/21
        Dim arr(0 To UBound(SelectedFields)) As Array
        'ReDim arr(0 To UBound(SelectedFields)) As Variant
        arr = SelectedFields
        rs = obj.GetMatch(obj, ConstructOrdrdFltrdQry(ConstructFilteredQuery(obj, arr), OrderBy))
        FltrdOrdrdMtch = GetResult(rs)
        obj.CloseDbs(obj)
        rs = Nothing
    End Function

    Public Function GetResult(rs As ADODB.Recordset) As Object
        If rs.RecordCount = 0 Then
            Return False
            Exit Function
        End If
        'Dim ResultArray() As Variant
        Dim ResultArray(0 To rs.RecordCount) As Array
        Dim i As Integer
        For i = 1 To rs.RecordCount
            ResultArray(i) = rs.GetRows(1)
        Next i
        Return ResultArray
    End Function

    Public Function ConstructMatchQuery(obj As aclsDataObject, Optional Filter As String = "*") As String
        Dim str As String
        If obj.Field1 = "" Then
            str = "SELECT " & Filter & " FROM " & obj.Db & " WHERE False"
            Return str
            Exit Function
        End If
        If VarType(obj.Value1) = vbString Then
            str = "SELECT " & Filter & " FROM " & obj.Db & " WHERE " & obj.Field1 & " = """ & obj.Value1 & """"
        End If

        If Not VarType(obj.Value1) = vbString Then
            str = "SELECT " & Filter & " FROM " & obj.Db & " WHERE " & obj.Field1 & " = " & obj.Value1 & ""
        End If
        If Not obj.Field2 = "" Then
            If VarType(obj.Value2) = vbString Then
                str = str & " AND " & obj.Field2 & " = """ & obj.Value2 & """"
            End If
            If Not VarType(obj.Value2) = vbString Then
                str = str & " AND " & obj.Field2 & " = " & obj.Value2 & ""
            End If
        End If
        Return str
    End Function

    Public Function ConstructFilteredQuery(obj As aclsDataObject, SelectedFields As Array) As String
        Dim str As String
        str = Join(SelectedFields(), ", ")
        Return ConstructMatchQuery(obj, str)
    End Function

    Public Function ConstructOrderedQuery(obj As aclsDataObject, OrderBy As String) As String

        Return ConstructMatchQuery(obj) & OrderBy
    End Function

    Public Function ConstructOrdrdFltrdQry(qry As String, OrderBy As String) As String

        Return qry & " ORDER BY " & OrderBy
    End Function
    '==========================================================================

    Public Sub Update(obj As aclsDataObject)
        Dim i As Integer
        rs = GetRecordsetMatch(obj, ConstructUpdateQuery(obj))

        If Not rs.RecordCount = 0 Then


            Do Until rs.EOF
                rs.Fields(obj.Field2).Value = obj.Value2
                rs.Update()
                rs.MoveNext()
            Loop
        End If
        obj.CloseDbs(obj)
        rs = Nothing
    End Sub

    Public Sub UpdateFromDict(obj As aclsDataObject, dict As Dictionary(Of String, Object))
        rs = GetRecordsetMatch(obj, ConstructMatchQuery(obj))
        If Not rs.RecordCount = 0 Then
            Dim fld As Object
            For Each fld In rs.Fields
                If Not fld.name = "ItemID" Then
                    If dict.ContainsKey(fld.name) Then

                        fld.value = dict(fld.name)
                    End If
                End If
            Next fld


            rs.Update()

        End If
        obj.CloseDbs(obj)
        rs = Nothing
    End Sub

    Public Function ConstructUpdateQuery(obj As aclsDataObject) As String
        Dim str As String
        If VarType(obj.Value1) = vbString Then
            str = "SELECT * FROM " & obj.Db & " WHERE " & obj.Field1 & " = """ & obj.Value1 & """"
        End If
        If Not VarType(obj.Value1) = vbString Then
            str = "SELECT * FROM " & obj.Db & " WHERE " & obj.Field1 & " = " & obj.Value1 & ""
        End If
        Return str
    End Function


    '==========================================================================




    'Public Function GetValueDict(obj As aclsDataObject, Optional qry As String = "") As Collection
    '    Dim coll As New Collection
    '    Dim dict As Scripting.Dictionary
    '    If Match(obj) = False Then

    '        dict = New Scripting.Dictionary
    '        coll.Add(dict)
    '        GetValueDict = coll
    '        dict = Nothing
    '        coll = Nothing
    '        Exit Function
    '    End If
    '    If qry = "" Then
    '        qry = ConstructMatchQuery(obj)
    '    End If
    '    rs = GetRecordsetMatch(obj, qry)
    '    Dim fld As Object
    '    Do Until rs.EOF
    '        dict = New Scripting.Dictionary
    '        For Each fld In rs.Fields
    '            'TODO - Reconfigure taxable field to be not calculated
    '            If Not fld.name = "Taxable" Then
    '                dict.Add(fld.name, fld.value)
    '            End If
    '        Next fld
    '        coll.Add(dict)
    '        rs.MoveNext()
    '    Loop
    '    GetValueDict = coll
    '    obj.CloseDbs(obj)
    '    coll = Nothing
    '    fld = Nothing
    '    dict = Nothing
    '    rs = Nothing
    'End Function


    Public Function GetValueDict(obj As aclsDataObject, Optional qry As String = "") As Collection
        Dim coll As New Collection
        Dim dict As Dictionary(Of String, Object)
        If Match(obj) = False Then

            dict = New Dictionary(Of String, Object)
            coll.Add(dict)
            GetValueDict = coll
            dict = Nothing
            coll = Nothing
            Exit Function
        End If
        If qry = "" Then
            qry = ConstructMatchQuery(obj)
        End If
        rs = GetRecordsetMatch(obj, qry)
        Dim fld As Object
        Do Until rs.EOF
            dict = New Dictionary(Of String, Object)
            For Each fld In rs.Fields
                'TODO - Reconfigure taxable field to be not calculated
                If Not fld.name = "Taxable" Then
                    dict.Add(fld.name, fld.value)
                End If
            Next fld
            coll.Add(dict)
            rs.MoveNext()
        Loop
        GetValueDict = coll
        obj.CloseDbs(obj)
        coll = Nothing
        fld = Nothing
        dict = Nothing
        rs = Nothing
    End Function


    '==========================================================================

    Public Sub AddNewRecord(obj As Object, dict As Dictionary(Of String, Object))

        Dim iDataObject As New aclsDataObject
        iDataObject = obj.Wrap(iDataObject)
        rs = GetRecordsetMatch(iDataObject, iDataObject.Db)

        rs.AddNew()
        On Error Resume Next
        For Each key As Object In dict.Keys
            rs.Fields(key).Value = dict(key)
        Next key
        rs.Update()
        On Error GoTo 0

        iDataObject.CloseDbs(iDataObject)

    End Sub

    '==========================================================================

    Public Sub ArchiveData(obj As aclsDataObject)
        Dim iDataObject As New aclsDataObject
        iDataObject.ArchiveData(obj, ConstructArchiveCmd(obj))
        iDataObject = Nothing
    End Sub

    Public Function ConstructArchiveCmd(obj As aclsDataObject) As String
        Dim str As String

        str = "INSERT INTO " & obj.Archive & " IN 'C:\Jared\POS\Access_Int\ReportsDB.accdb' SELECT * FROM " & obj.Db & ""

        Return str
    End Function

End Module
