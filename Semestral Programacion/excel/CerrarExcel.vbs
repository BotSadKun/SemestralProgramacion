Set objExcel = GetObject(, "Excel.Application")
If Not objExcel Is Nothing Then
    objExcel.Quit
End If
