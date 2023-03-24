Public Class DateCompte : Implements IComparable(Of DateCompte), IEquatable(Of DateCompte)
    Private ReadOnly aDay As Integer
    Private ReadOnly aMonth As Integer
    Private ReadOnly aYear As Integer
    Private ReadOnly aText As String
    Public Sub New(textValue As String)
        If textValue IsNot Nothing AndAlso textValue.Split(".").Count = 3 Then
            Dim ThisDate As String() = textValue.Split(".")
            aYear = CInt(ThisDate(2))
            aMonth = CInt(ThisDate(1))
            aDay = CInt(ThisDate(0))
        End If
        aText = textValue
    End Sub
    Public ReadOnly Property Day As Integer
        Get
            Return aDay
        End Get
    End Property
    Public ReadOnly Property Mounth As Integer
        Get
            Return aMonth
        End Get
    End Property
    Public ReadOnly Property Year As Integer
        Get
            Return aYear
        End Get
    End Property
    Public ReadOnly Property AsText As String
        Get
            Return aText
        End Get
    End Property
    Public Function CompareTo(other As DateCompte) As Integer Implements IComparable(Of DateCompte).CompareTo
        If other Is Nothing Then
            Return String.Compare(aText, Nothing, StringComparison.CurrentCulture)
        Else
            If aYear <> other.aYear Then
                Return aYear - other.aYear
            ElseIf aMonth <> other.aMonth Then
                Return aMonth - other.aMonth
            Else
                Return aDay - other.aDay
            End If
        End If
    End Function
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Me.GetType() IsNot obj.GetType() Then
            Return False
        End If
        Return Me.Equals(CType(obj, DateCompte))
    End Function
    Public Overloads Function Equals(other As DateCompte) As Boolean Implements IEquatable(Of DateCompte).Equals
        Return CompareTo(other)
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return (aDay, aMonth, aYear).GetHashCode()
    End Function
    Public Shared Operator =(left As DateCompte, right As DateCompte) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) = 0
        Else
            Return right Is Nothing
        End If
    End Operator
    Public Shared Operator <>(left As DateCompte, right As DateCompte) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) <> 0
        Else
            Return right IsNot Nothing
        End If
    End Operator
    Public Shared Operator >(left As DateCompte, right As DateCompte) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) > 0
        Else
            Return False
        End If
    End Operator
    Public Shared Operator <(left As DateCompte, right As DateCompte) As Boolean
        Return Not left >= right
    End Operator
    Public Shared Operator >=(left As DateCompte, right As DateCompte) As Boolean
        If left IsNot Nothing Then
            Return left.CompareTo(right) >= 0
        Else
            Return False
        End If
    End Operator
    Public Shared Operator <=(left As DateCompte, right As DateCompte) As Boolean
        Return Not left > right
    End Operator
End Class
