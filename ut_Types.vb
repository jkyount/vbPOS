Module ut_Types
    Public Function ReflectType(ClassName As String) As Object
        Return Activator.CreateInstance(vbNullString, ThisAssembly & "." & ClassName).Unwrap
    End Function

End Module
