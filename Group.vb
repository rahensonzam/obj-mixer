Public NotInheritable Class Group
    Private _name As String
    Private _faces As New List(Of Integer)

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Faces() As List(Of Integer)
        Get
            Return _faces
        End Get
        Set(ByVal value As List(Of Integer))
            _faces = value
        End Set
    End Property

    Public Sub New(ByVal name As String)
        _name = name
    End Sub

End Class
