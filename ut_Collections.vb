Imports Scripting
Imports System.Collections.ObjectModel
Imports System.Windows


Module ut_Collections
    Public Sub ClearCollection(coll As Collection)
        If Not coll.Count = 0 Then
            For i As Integer = 1 To coll.Count
                coll.Remove(1)
            Next i
        End If
    End Sub

    Public Function DuplicateCollection(coll As Collection) As Collection
        Dim NewColl As New Collection
        For i As Integer = 1 To coll.Count
            NewColl.Add(coll(i))
        Next i
        DuplicateCollection = NewColl
        If coll.Count = 0 Then
            MsgBox("Attempted to duplicate an empty collection.")
        End If
    End Function

    Public Function DuplicateCheckLines(coll As Collection) As Collection
        Dim NewColl As New Collection
        For i As Integer = 1 To coll.Count
            NewColl.Add(coll(i), "Line" & i)
        Next i
        DuplicateCheckLines = NewColl
        If coll.Count = 0 Then
            MsgBox("Attempted to duplicate an empty collection.")
        End If
    End Function

    Public Function DuplicateCItem(coll As Collection) As Collection
        Dim NewColl As New Collection
        For i As Integer = 1 To coll.Count
            NewColl.Add(coll(i), CStr(coll(i).CollID))
        Next i
        DuplicateCItem = NewColl
    End Function

    Public Function RecallCheckLines(check As String) As Collection


        Dim x As New zclsDailyCheckDetail
        Dim OrderBy As String
        OrderBy = "ORDER BY Seat ASC, LocalGroup ASC"
        Dim iDataObj As New aclsDataObject
        iDataObj = x.Wrap(GetNewMatchObj(, check))
        RecallCheckLines = SortCheckLines(GetCheckLines(check, GetValueDict(iDataObj, ConstructOrderedQuery(iDataObj, OrderBy))))

        x = Nothing
        iDataObj = Nothing

    End Function



    Public Function SortCheckLines(coll As Collection) As Collection
        Dim collTemp As New Collection
        Dim z As zclsCheckLines
        Dim k As Integer = 1
        If Not coll.Count = 0 Then
            For i As Integer = 1 To 12
                For Each z In coll
                    If z.Seat = i Then
                        collTemp.Add(z, ("Line" & k))
                        z.Row = k
                        k += 1
                    End If
                Next z
            Next i

            Dim st As Integer = collTemp("Line1").seat

            Dim GuiRow As Integer = 2
            For i As Integer = 1 To collTemp.Count
                If Not collTemp("Line" & i).seat = st Then
                    st = collTemp("Line" & i).seat
                    GuiRow += GuiRow
                End If
                collTemp("Line" & i).GuiRow = GuiRow
                GuiRow += GuiRow
            Next i
        End If
EH:
        SortCheckLines = collTemp
        coll = Nothing
        collTemp = Nothing
    End Function

    Public Function NullToZero(arr As Object) As Object
        For i As Integer = 1 To UBound(arr)
            If IsDBNull(arr(i)) Then arr(i) = CStr(0)
        Next i
        NullToZero = arr
    End Function

    Public Sub ReplaceDictValue(dict As Dictionary(Of String, Object), key As String, NewVal As Object)
        For Each member As Object In dict.Keys
            If member = key Then
                dict.Remove(key)
                dict.Add(key, NewVal)
                Exit Sub
            End If
        Next member
    End Sub

    Public Function OrderByCollID(coll As Collection) As Collection
        If coll.Count = 1 Then
            OrderByCollID = coll
            Exit Function
        End If

        Dim NewColl As Collection
        NewColl = New Collection

        Do Until coll.Count = 0
            For Each member As aclsItem In coll
                For i As Integer = 1 To coll.Count
                    If Not member.CollID <= coll(i).CollID Then
                        GoTo NextMember
                    End If
                Next i
                NewColl.Add(member, CStr(member.CollID))

                coll.Remove(CStr(member.CollID))

NextMember:
            Next member
        Loop
        OrderByCollID = NewColl
    End Function

    Public Function GetNextCollID(coll As Collection) As Integer
        If Not coll.Count = 0 Then
            GetNextCollID = coll(coll.Count).CollID + 1
            Exit Function
        End If
        GetNextCollID = 1
    End Function

    Public Function OrderByParent(coll As Collection) As Collection
        If coll.Count = 1 Then
            OrderByParent = coll
            Exit Function
        End If

        Dim NewColl As New Collection
        NewColl.Add(coll(1), CStr(coll(1).CollID))

        For Each member As aclsItem In coll
            If member.Parent.ID = coll(1).CollID Then
                NewColl.Add(member, CStr(member.CollID))
                AddChildrenToColl(NewColl, member)
            End If
        Next member
        OrderByParent = NewColl
    End Function


    Public Sub AddChildrenToColl(coll As Collection, member As aclsItem)

        Dim item As aclsItem
        For Each child As bclsChild In member.Children.coll
            item = GetItemByID(child.ID)
            coll.Add(item, CStr(item.CollID))
            AddChildrenToColl(coll, item)
        Next child

    End Sub

    Public Function CObservableColl(coll As IEnumerable(Of Object)) As ObservableCollection(Of Object)
        Return New ObservableCollection(Of Object)(coll)
    End Function



    Public Sub CloseWindow(Name As String)
        Dim coll As Collection = GetWindows()
        coll(Name).Close()

    End Sub

    Public Sub OpenWindow(Wndw As Window)
        Wndw.Activate()
    End Sub
End Module
