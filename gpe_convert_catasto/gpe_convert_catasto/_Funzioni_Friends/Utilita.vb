
Module Utilita

    Public Function NoNullText(ByVal Value As String) As Object
        Return If(Value.Trim = "", DBNull.Value, Value).ToString.Trim
    End Function

    Public Function NoNullDate(ByVal Value As String) As Object
        Return If(IsDate(Value.Trim) = False, DBNull.Value, Convert.ToDateTime(Value))
    End Function

    Public Function NoNullInteger(ByVal Value As String) As Integer
        Return If(Val(Value) = 0, 0, Convert.ToInt32(Value))
    End Function

    Public Function NoNullDouble(ByVal Value As String) As Double
        Return If(Val(Value) = 0, 0, Convert.ToDouble(Value))
    End Function

    Friend Function ApiciSI(ByVal Testo As Object) As Object
        Return If(Testo Is Nothing, DBNull.Value, Chr(34) & (Testo.Trim).Replace(Chr(34), "") & Chr(34))
    End Function

    Friend Function DbValue(ByVal ValueText As String) As Object
        Return If(String.IsNullOrEmpty(ValueText) = True, DBNull.Value, ValueText)

    End Function
End Module
