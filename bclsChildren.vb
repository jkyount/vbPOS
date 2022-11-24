Public Class bclsChildren
    Private pColl As New Collection


    Public Property coll As Collection
        Set(value As Collection)
            pColl = value
        End Set
        Get
            'TODO IMPLEMENT coll AS DEFAULT CLASS MEMBER, AS BELOW
            'Attribute coll.VB_UserMemId = 0
            Return pColl
        End Get
    End Property

    Public Sub New()

    End Sub

    'Public Function GetNew() As bclsChildren
    '    Dim x As New bclsChildren
    '    GetNew = x
    '    x = Nothing
    'End Function




End Class
