Imports System
Imports MyFirstAddIn.FirstExcelDnaAddIn_m
Imports System.Text

Module Program
	Sub Main(args As String())

		Dim q

		Dim Arr1() As Int64 = {1, 2, 3}
		SetTranspArr("q", Arr1)
		q = GetArr("q")

		Dim Arr2(,) As Double = {{1}, {2}, {3}}
		SetTranspArr("q", Arr2)
		q = GetArr("q")

		Dim Arr3(,) As Byte = {{1, 2, 3}, {4, 5, 6}}
		SetTranspArr("q", Arr3)
		q = GetArr("q")

		Dim Arr4(,) = {{"one"}, {"tw"}, {"tree"}, {"four"}}
		SetTranspArr("q", Arr4)
		q = GetArr("q")

		Dim Arr5() = {"one", "tw", "tree", "four"}
		SetTranspArr("q", Arr5)
		q = GetArr("q")

		Dim Arr(,) = {{"one", "tw"}, {"tree", "four"}}
		SetTranspArr("q", Arr)
		q = GetArr("q")
		q = GetArr("q", 1, 2)


		Dim a(,) As Integer = {{1, 2, 3}, {4, 5, 6}}
		'dim t = a.GetType()

		Dim oA
		oA = a

		'1
		oA(0, 0) = 0

		'2
		CType(oA, Array)(0, 1) = 0

		'3
		Dim aa = CType(oA, Array)
		aa(0, 2) = 0


		Dim r = oA.Rank()

		Dim b() = New Integer(6) {}

		oA = b
		r = CType(oA, Array).Rank

		Buffer.BlockCopy(a, 3 * Len(a(1, 1)), b, 1 * Len(a(1, 1)), 3 * Len(a(1, 1)))

		Array.Copy(a, 3, b, 4, 3)

	End Sub
End Module
