Public Class BookLine : Implements IComparable(Of BookLine), IEquatable(Of BookLine)

    Private Cptegen As String
    Private Rubrique As String
    Private NumeroFlux As String
    Private Nom As String
    Private Libelle As String
    Private MntEngHTR As Double
    Private MontantPa As Double
    Private Rapprochmt As String
    Private RefFactF As String
    Private DatePce As String
    Private DCompt As DateCompte
    Private NumPiece As String
    Private Comment As String
    Private From As String

    Public Property ACptegen As String
        Get
            Return Cptegen
        End Get
        Set(value As String)
            Cptegen = value
        End Set
    End Property
    Public Property BRubrique As String
        Get
            Return Rubrique
        End Get
        Set(value As String)
            Rubrique = value
        End Set
    End Property
    Public Property CNumeroFlux As String
        Get
            Return NumeroFlux
        End Get
        Set(value As String)
            NumeroFlux = value
        End Set
    End Property
    Public Property DNom As String
        Get
            Return Nom
        End Get
        Set(value As String)
            Nom = value
        End Set
    End Property
    Public Property ELibelle As String
        Get
            Return Libelle
        End Get
        Set(value As String)
            Libelle = value
        End Set
    End Property
    Public Property FMntEngHtr As Double
        Get
            Return MntEngHTR
        End Get
        Set(value As Double)
            MntEngHTR = value
        End Set
    End Property
    Public Property GMontantPA As Double
        Get
            Return MontantPa
        End Get
        Set(value As Double)
            MontantPa = value
        End Set
    End Property
    Public Property HRapprochmt As String
        Get
            Return Rapprochmt
        End Get
        Set(value As String)
            Rapprochmt = value
        End Set
    End Property
    Public Property IRefFactF As String
        Get
            Return RefFactF
        End Get
        Set(value As String)
            RefFactF = value
        End Set
    End Property
    Public Property JDatePce As String
        Get
            Return DatePce
        End Get
        Set(value As String)
            DatePce = value
        End Set
    End Property
    Public Property KDCompt As DateCompte
        Get
            Return DCompt
        End Get
        Set(value As DateCompte)
            DCompt = value
        End Set
    End Property
    Public Property LNumPiece As String
        Get
            Return NumPiece
        End Get
        Set(value As String)
            NumPiece = value
        End Set
    End Property
    Public Property MComment As String
        Get
            Return Comment
        End Get
        Set(value As String)
            Comment = value
        End Set
    End Property
    Public Property NFrom As String
        Get
            Return From
        End Get
        Set(value As String)
            From = value
        End Set
    End Property

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Me.GetType() IsNot obj.GetType() Then
            Return False
        End If
        Return Me.Equals(CType(obj, BookLine))
    End Function

    Public Overloads Function Equals(other As BookLine) As Boolean Implements IEquatable(Of BookLine).Equals
        Return other IsNot Nothing AndAlso
               BRubrique = other.BRubrique AndAlso
               CNumeroFlux = other.CNumeroFlux AndAlso
               ELibelle = other.ELibelle AndAlso
               FMntEngHtr = other.FMntEngHtr AndAlso
               JDatePce = other.JDatePce
    End Function

    Public Function CompareTo(other As BookLine) As Integer Implements IComparable(Of BookLine).CompareTo
        If other IsNot Nothing Then
            If KDCompt IsNot Nothing Then
                Return KDCompt.CompareTo(other.KDCompt)
            ElseIf other.KDCompt IsNot Nothing Then
                Return other.KDCompt.CompareTo(KDCompt)
            Else
                Return 0
            End If
        Else
            Return 1
        End If
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return (ACptegen, BRubrique, CNumeroFlux, DNom, ELibelle, FMntEngHtr, GMontantPA, HRapprochmt, IRefFactF, JDatePce, KDCompt, LNumPiece, MComment, NFrom).GetHashCode()
    End Function

    Public Shared Operator =(left As BookLine, right As BookLine) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) = 0
        Else
            Return right Is Nothing
        End If
    End Operator
    Public Shared Operator <>(left As BookLine, right As BookLine) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) <> 0
        Else
            Return right IsNot Nothing
        End If
    End Operator
    Public Shared Operator >(left As BookLine, right As BookLine) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) > 0
        Else
            Return False
        End If
    End Operator

    Public Shared Operator <(left As BookLine, right As BookLine) As Boolean
        Return Not left >= right
    End Operator

    Public Shared Operator >=(left As BookLine, right As BookLine) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) >= 0
        Else
            Return False
        End If
    End Operator

    Public Shared Operator <=(left As BookLine, right As BookLine) As Boolean
        Return Not left > right
    End Operator

    Public Overrides Function ToString() As String
        Dim asText As String = ""
        asText &= " " & Cptegen
        asText &= " " & Rubrique
        asText &= " " & NumeroFlux
        asText &= " " & Nom
        asText &= " " & Libelle
        asText &= " " & MntEngHTR
        asText &= " " & MontantPa
        asText &= " " & Rapprochmt
        asText &= " " & RefFactF
        asText &= " " & DatePce
        asText &= " " & "DCompt"
        asText &= " " & NumPiece
        asText &= " " & Comment
        asText &= " " & From
        Return asText
    End Function
End Class