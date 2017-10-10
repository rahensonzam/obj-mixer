Public NotInheritable Class Face

    Private _faceVertices As List(Of Integer())
    Private _id As Integer
    Private _newId As Integer

    Public Property FaceVertices() As List(Of Integer())
        Get
            Return _faceVertices
        End Get
        Set(ByVal value As List(Of Integer()))
            _faceVertices = value
        End Set
    End Property

    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property NewId() As Integer
        Get
            Return _newId
        End Get
        Set(ByVal value As Integer)
            _newId = value
        End Set
    End Property

    Public Sub New(ByVal faceVertices As List(Of Integer()), ByVal id As Integer)
        _faceVertices = faceVertices
        _id = id
        _newId = id
    End Sub

End Class
