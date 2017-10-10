Imports obj_mixer.frmMain

Public NotInheritable Class Vertex

    Private _pos As loc3d
    Private _norm As loc3d
    Private _tex As loc2d
    Private _id As Integer
    Private _newId As Integer

    Public Property Pos() As loc3d
        Get
            Return _pos
        End Get
        Set(ByVal value As loc3d)
            _pos = value
        End Set
    End Property

    Public Property Norm() As loc3d
        Get
            Return _norm
        End Get
        Set(ByVal value As loc3d)
            _norm = value
        End Set
    End Property

    Public Property Tex() As loc2d
        Get
            Return _tex
        End Get
        Set(ByVal value As loc2d)
            _tex = value
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

    Public Sub New(ByVal id As Integer)
        _id = id
        _newId = id
    End Sub

End Class
