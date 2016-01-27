Public Module LazyProperty
    Public Function SetValueIfEmpty(Of T As Class)(ByRef Varible As T, Calculate As Func(Of T)) As T
        If Varible Is Nothing Then Varible = Calculate.Invoke
        Return Varible
    End Function
    Public Function SetValueIfEmpty(Of T As Structure)(ByRef Varible As T?, Calculate As Func(Of T)) As T
        If Varible Is Nothing Then Varible = Calculate.Invoke
        Return Varible.Value
    End Function
End Module
