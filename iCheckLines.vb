Imports System.Collections.ObjectModel

Module iCheckLines
    Public Function GetCheckLines(check As String, ValueDict As Collection) As Collection
        Dim iCheckLines As New zclsCheckLines
        GetCheckLines = iCheckLines.GetCheckLines(check, ValueDict)
        iCheckLines = Nothing
    End Function
End Module
