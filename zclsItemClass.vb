
Public Class zclsItemClass

    '// Obtains an instance of a type derived from BaseItemClass,
    '// from a predefined list of types.  Within BaseItemClass's derived types
    '// there must exist a type whose name matches the "Class Name" field of a 
    '// record contained in the database to which zclsItemClass connects.  Only types
    '// explicitly defined in zclsItemClass.GetClassColl() are available for instantiation.


    Inherits DBObject

    Public dbItemClass As New ADODB.Connection
    Public rsItemClass As New ADODB.Recordset


    Public Overrides Function GetDb() As String
        GetDb = "ItemClass"
    End Function

    Public Overrides Function GetDbFile() As String
        GetDbFile = "ItemClass"
    End Function

    Public Overrides Function GetConn() As ADODB.Connection
        GetConn = dbItemClass
    End Function

    Public Overrides Function GetRs() As ADODB.Recordset
        GetRs = rsItemClass
    End Function

    Public Function GetItemType(ClassCode As String) As BaseItemClass
        Dim CCode As Integer = CInt(ClassCode)
        Dim ClassName As String = ValueMatch(Wrap(GetNewMatchObj("ClassCode", CCode)), "ClassType")

        If CCode = 1 Then
            If CItem.Count > 0 Then
                ClassName = ValueMatch(Wrap(GetNewMatchObj("ClassCode", 2)), "ClassType")
            End If
        End If
        Return ReflectType(ClassName)
    End Function

    'Public Function GetItemType(ClassCode As String) As BaseItemClass
    '    Dim coll As Collection = GetClassColl()
    '    GetItemType = coll(ClassCode)
    '    If ClassCode = 1 Then
    '        If CItem.Count > 0 Then
    '            GetItemType = coll(2)
    '        End If
    '    End If
    '    coll = Nothing
    'End Function

    Public Function GetClassColl() As Collection
        Dim ClassColl As New Collection
        Dim x As Object
        x = New zclsPrimary
        ClassColl.Add(x, GetClassCode(TypeName(x)))
        x = New zclsComponent
        ClassColl.Add(x, GetClassCode(TypeName(x)))
        'x = New zclsCustomItem
        'ClassColl.Add(x, GetClassCode(TypeName(x)))
        'x = New zclsSpecialInstruction
        'ClassColl.Add(x, GetClassCode(TypeName(x)))
        'x = New zclsDiscount
        'ClassColl.Add(x, GetClassCode(TypeName(x)))
        'x = New zclsPizza
        'ClassColl.Add(x, GetClassCode(TypeName(x)))
        'x = New zclsPizzaTopping
        'ClassColl.Add(x, GetClassCode(TypeName(x)))
        'x = New zclsMod
        'ClassColl.Add(x, GetClassCode(TypeName(x)))
        x = Nothing
        GetClassColl = ClassColl
    End Function

    Public Function GetClassCode(ClassType As String) As String
        OpenDbs()
        rsItemClass.let_Source("SELECT * FROM ItemClass WHERE ClassType = """ & ClassType & """")
        rsItemClass.Open()
        GetClassCode = CStr(rsItemClass.Fields("ClassCode").Value)
        rsItemClass.Close()
        dbItemClass.Close()
    End Function

    Public Function GetAssignableClassNames() As Object
        Dim ClassNameArray As Object
        Dim ClassCodeArray As Object

        ClassNameArray = FilteredMatch(Wrap(GetNewMatchObj("Assignable", True)), {"ClassName"})
        ClassCodeArray = FilteredMatch(Wrap(GetNewMatchObj("Assignable", True)), {"ClassCode"})

        Dim i As Integer
        For i = 1 To UBound(ClassNameArray)
            ClassCodeArray(i)(0, 0) = ClassCodeArray(i)(0, 0) & " - " & ClassNameArray(i)(0, 0)
        Next i
        GetAssignableClassNames = ClassCodeArray
    End Function
End Class
