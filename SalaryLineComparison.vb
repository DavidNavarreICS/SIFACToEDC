Public Class SalaryLineComparison : Implements IComparer(Of BookLine)
    Public Function Compare(x As BookLine, y As BookLine) As Integer Implements IComparer(Of BookLine).Compare
        If x IsNot Nothing Then
            If y IsNot Nothing Then
                Return String.CompareOrdinal(x.DNom, y.DNom)
            Else
                Return String.CompareOrdinal(x.DNom, Nothing)
            End If
        Else
            If y IsNot Nothing Then
                Return String.CompareOrdinal(Nothing, y.DNom)
            Else
                Return 0
            End If
        End If
    End Function
End Class
