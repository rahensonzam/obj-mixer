Public NotInheritable Class Model

    Private _name As String
    Private _vertices As New Dictionary(Of Integer, Vertex)
    Private _faces As New Dictionary(Of Integer, Face)
    Private _groups As New Dictionary(Of String, Group)
    Private _selectedGroups As New List(Of String)

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Vertices() As Dictionary(Of Integer, Vertex)
        Get
            Return _vertices
        End Get
        Set(ByVal value As Dictionary(Of Integer, Vertex))
            _vertices = value
        End Set
    End Property

    Public Property Faces() As Dictionary(Of Integer, Face)
        Get
            Return _faces
        End Get
        Set(ByVal value As Dictionary(Of Integer, Face))
            _faces = value
        End Set
    End Property

    Public Property Groups() As Dictionary(Of String, Group)
        Get
            Return _groups
        End Get
        Set(ByVal value As Dictionary(Of String, Group))
            _groups = value
        End Set
    End Property

    Public Property SelectedGroups() As List(Of String)
        Get
            Return _selectedGroups
        End Get
        Set(ByVal value As List(Of String))
            _selectedGroups = value
        End Set
    End Property

    Public Sub New(ByVal name As String, ByVal vertices As Dictionary(Of Integer, Vertex), ByVal faces As Dictionary(Of Integer, Face), ByVal groups As Dictionary(Of String, Group))
        _name = name
        _vertices = vertices
        _faces = faces
        _groups = groups
    End Sub

End Class
