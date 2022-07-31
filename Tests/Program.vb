Imports System
Imports MyFirstAddIn.FirsExcelDnaAddIn_m

Module Program
	Sub Main(args As String())
		SetArr("q", {{1, 2}, {3, 4}})
		Dim a = GetArrRow("q", 2)
	End Sub
End Module
