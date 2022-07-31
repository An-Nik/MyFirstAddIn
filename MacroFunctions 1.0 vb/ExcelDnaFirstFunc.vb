Imports System
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Linq

Public Module FirstExcelDnaAddIn_m

#Region "  Call UDF"

	Public DictOfWSheetObjects As New Dictionary(Of String, Object)
	Const _shName = "[TblWork_1.2.xlam]"

	<ExcelFunction(Description:="Вызвать указанную UDF, определённую в той же книге, методом Run. Имя книги берётся из App.Caller", Category:="ANik")>
	Public Function RunFunc_pa(FuncName As String, ParamArray Arr() As Object)

		'получить имя книги используя Application.Caller
		Dim app = CType(ExcelDnaUtil.Application, Excel.Application)
		Dim callerCell = CType(app.Caller, Excel.Range)
		Dim callerCellAddr As String = callerCell.Address(,,, True)
		Dim wbName As String = Replace(Replace(Mid(callerCellAddr, 1, InStr(callerCellAddr, "]") - 1), "[", ""), "'", "")
		Dim wbFullName As String = "'" + wbName + "'!" + FuncName

		Return app.Run(wbFullName, Arr)

	End Function


	<ExcelFunction(Description:="Вызвать указанную UDF, определённую в той же книге, методом Run. Имя книги берётся из App.Caller", Category:="ANik")>
	Public Function RunFunc(FuncName As String, Optional par1 As Object = Nothing, Optional par2 As Object = Nothing,
										Optional par3 As Object = Nothing, Optional par4 As Object = Nothing, Optional par5 As Object = Nothing)

		'получить имя книги используя Application.Caller
		Dim app = CType(ExcelDnaUtil.Application, Excel.Application)
		Dim callerCell = CType(app.Caller, Excel.Range)
		Dim callerCellAddr As String = callerCell.Address(,,, True)
		Dim wbName As String = Replace(Replace(Mid(callerCellAddr, 1, InStr(callerCellAddr, "]") - 1), "[", ""), "'", "")
		Dim wbFullName As String = "'" + wbName + "'!" + FuncName

		If TypeOf par1 Is ExcelMissing Then
			Return app.Run(wbFullName)
		ElseIf TypeOf par2 Is ExcelMissing Then
			Return app.Run(wbFullName, par1)
		ElseIf TypeOf par3 Is ExcelMissing Then
			Return app.Run(wbFullName, par1, par2)
		ElseIf TypeOf par4 Is ExcelMissing Then
			Return app.Run(wbFullName, par1, par2, par3)
		ElseIf TypeOf par5 Is ExcelMissing Then
			Return app.Run(wbFullName, par1, par2, par3, par4)
		Else
			Return app.Run(wbFullName, par1, par2, par3, par4, par5)
		End If

	End Function


	<ExcelFunction(Description:="Вызвать указанную UDF либо из надстройки [TblWork_1.2.xlam] (тогда shName=''), либо из модуля рабочего листа (тогда shName='[имяКниги.расш]ИмяЛиста')", Category:="ANik")>
	Public Function CallBN(shName As String, FuncName As String, Optional par1 As Object = Nothing, Optional par2 As Object = Nothing,
										Optional par3 As Object = Nothing, Optional par4 As Object = Nothing, Optional par5 As Object = Nothing)

		Dim app = CType(ExcelDnaUtil.Application, Excel.Application)
		Dim wsObject As Object  'указывает либо на надстройку, либо на книгу, откуда пришёл вызов
		Dim wbName As String

		'•определить, откуда требуется вызвать функцию: из надстройки или с листа, на котором она расположена
		' если имя листа задано
		If shName <> "" Then

			'вызвать функцию с листа, откуда пришёл вызов
			'--------------------------------------------

			'проверить наличие в словаре ссылки на лист, с которого пришёл вызов
			If DictOfWSheetObjects.TryGetValue(shName, wsObject) = False Then

				'в словаре ссылки на кнмгу/лист нет, добавить её
				'-----------------------------------------------

				'выделить из параметра shName название книги и листа
				wbName = Mid(shName, 2, InStr(shName, "]") - 2)
				Dim wsName = Mid(shName, InStr(shName, "]") + 1, 100)
				wsObject = CType(app.Workbooks(wbName).Worksheets(wsName), Excel.Worksheet)

				'добавить в словарь лист, откуда пришёл вызов
				DictOfWSheetObjects(shName) = wsObject

			Else
				'в словаре ссылка на лист есть, получена в wsObject
			End If

		Else
			'имя листа не задано, вызвать функции из надстройки
			'--------------------------------------------------

			'проверить наличие ключа c именем надстройки в словаре
			If DictOfWSheetObjects.TryGetValue(_shName, wsObject) = False Then

				'в словаре ссылки на надстройку нет, добавить её
				'-----------------------------------------------

				'выделить из параметра shName название книги и листа
				wbName = Mid(_shName, 2, InStr(_shName, "]") - 2)
				'wsName = Mid(_shName, InStr(_shName, "]") + 1, 100)
				wsObject = app.Workbooks(wbName)

				'добавить в словарь лист, откуда пришёл вызов
				DictOfWSheetObjects(_shName) = wsObject

			Else
				'в словаре ссылка на надстройку есть, получена в wsObject
			End If

		End If

		'вызвать пользовательскую функцию
		'--------------------------------
		If TypeOf par1 Is ExcelMissing Then
			Return CallByName(wsObject, FuncName, vbMethod)
		ElseIf TypeOf par2 Is ExcelMissing Then
			Return CallByName(wsObject, FuncName, vbMethod, par1)
		ElseIf TypeOf par3 Is ExcelMissing Then
			Return CallByName(wsObject, FuncName, vbMethod, par1, par2)
		ElseIf TypeOf par4 Is ExcelMissing Then
			Return CallByName(wsObject, FuncName, vbMethod, par1, par2, par3)
		ElseIf TypeOf par5 Is ExcelMissing Then
			Return CallByName(wsObject, FuncName, vbMethod, par1, par2, par3, par4)
		Else
			Return CallByName(wsObject, FuncName, vbMethod, par1, par2, par3, par4, par5)
		End If

	End Function

#End Region


#Region "  SetVar / GetVar / ClearVarDict "
	'--------------------------------------------------------------------------------------------------------------------

	Private DictOfVars As New Dictionary(Of String, Object)

	<ExcelFunction(Description:="Задать значение переменной", Category:="ANik")>
	Public Function SetVar(<ExcelArgument(Description:="Имя переменной")> VarName As String,
									  <ExcelArgument(Description:="Значение")> VarValue As Object,
									  <ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)

		DictOfVars(VarName) = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue

	End Function


	<ExcelFunction(Description:="Получить значение переменной по имени", Category:="ANik")>
	Public Function GetVar(<ExcelArgument(Description:="Имя переменной")> VarName As String)

		Dim VarValue

		If Not DictOfVars.TryGetValue(VarName, VarValue) Then
			Return Excel.XlCVError.xlErrValue
		Else
			Return VarValue
		End If

	End Function


	<ExcelFunction(Description:="Удалить все именные переменные", Category:="ANik")>
	Public Function ClearVarDict()

		DictOfVars.Clear()
		Return 0

	End Function

#End Region


#Region "  SetIntN / GetIntN "
	'--------------------------------------------------------------------------------------------------------------------
	Private Int1 As Long
	<ExcelFunction(Description:="Задать значение целочисленной переменной", Category:="ANik")>
	Public Function SetInt1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Int1 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Int2 As Long
	<ExcelFunction(Description:="Задать значение целочисленной переменной", Category:="ANik")>
	Public Function SetInt2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Int2 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Int3 As Long
	<ExcelFunction(Description:="Задать значение целочисленной переменной", Category:="ANik")>
	Public Function SetInt3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Int3 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Int4 As Long
	<ExcelFunction(Description:="Задать значение целочисленной переменной", Category:="ANik")>
	Public Function SetInt4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Int4 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Int5 As Long
	<ExcelFunction(Description:="Задать значение целочисленной переменной", Category:="ANik")>
	Public Function SetInt5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Int5 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function
	'--------------------------------------------------------------------------------------------------------------------


	<ExcelFunction(Description:="Получить сохранённое значение целочисленной переменной", Category:="ANik")>
	Public Function GetInt1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Int1
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение целочисленной переменной", Category:="ANik")>
	Public Function GetInt2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Int2
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение целочисленной переменной", Category:="ANik")>
	Public Function GetInt3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Int3
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение целочисленной переменной", Category:="ANik")>
	Public Function GetInt4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Int4
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение целочисленной переменной", Category:="ANik")>
	Public Function GetInt5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Int5
	End Function

#End Region


#Region "  SetDoubN / GetDoubN "
	'--------------------------------------------------------------------------------------------------------------------
	Private Doub1 As Double
	<ExcelFunction(Description:="Задать значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function SetDoub1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Double = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Doub1 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Doub2 As Double
	<ExcelFunction(Description:="Задать значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function SetDoub2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Double = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Doub2 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Doub3 As Double
	<ExcelFunction(Description:="Задать значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function SetDoub3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Double = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Doub3 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Doub4 As Double
	<ExcelFunction(Description:="Задать значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function SetDoub4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Double = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Doub4 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Doub5 As Double
	<ExcelFunction(Description:="Задать значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function SetDoub5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Double = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Doub5 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function
	'--------------------------------------------------------------------------------------------------------------------


	<ExcelFunction(Description:="Получить сохранённое значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function GetDoub1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Doub1
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function GetDoub2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Doub2
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function GetDoub3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Doub3
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function GetDoub4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Doub4
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение дробной (вещественной) переменной", Category:="ANik")>
	Public Function GetDoub5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Doub5
	End Function

#End Region


#Region "  SetStrN / GetStrN "
	'--------------------------------------------------------------------------------------------------------------------

	Private Str1 As String
	<ExcelFunction(Description:="Задать значение текстовой переменной", Category:="ANik")>
	Public Function SetStr1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Str1 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Str2 As String
	<ExcelFunction(Description:="Задать значение текстовой переменной", Category:="ANik")>
	Public Function SetStr2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Str2 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Str3 As String
	<ExcelFunction(Description:="Задать значение текстовой переменной", Category:="ANik")>
	Public Function SetStr3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Str3 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Str4 As String
	<ExcelFunction(Description:="Задать значение текстовой переменной", Category:="ANik")>
	Public Function SetStr4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Str4 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Str5 As String
	<ExcelFunction(Description:="Задать значение текстовой переменной", Category:="ANik")>
	Public Function SetStr5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Значение")> Optional VarValue As Long = 0,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Str5 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function
	'--------------------------------------------------------------------------------------------------------------------


	<ExcelFunction(Description:="Получить сохранённое значение текстовой переменной", Category:="ANik")>
	Public Function GetStr1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Str1
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение текстовой переменной", Category:="ANik")>
	Public Function GetStr2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Str2
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение текстовой переменной", Category:="ANik")>
	Public Function GetStr3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Str3
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение текстовой переменной", Category:="ANik")>
	Public Function GetStr4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Str4
	End Function

	<ExcelFunction(Description:="Получить сохранённое значение текстовой переменной", Category:="ANik")>
	Public Function GetStr5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "")
		Return Str5
	End Function

#End Region


#Region "  SetArr/SetArrN, GetArr/GetArrN, GetArrRow "
	'==================================================

	Private DictOfArrays As New Dictionary(Of String, Object)

	'--------------------------------------------------------------------------------------------------------------------
	<ExcelFunction(Description:="Сохранить имя и значения массива в словаре", Category:="ANik")
	>'-------------------------------------------------------------------------------------------------------------------
	Public Function SetArr(<ExcelArgument(Description:="Имя сохраняемого массива")> arrName As String,
									<ExcelArgument(Description:="Массив")> Arr As Object,
									<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		DictOfArrays(arrName) = Arr
		If TypeOf ReturnValue Is ExcelMissing Then Return Arr Else Return ReturnValue
	End Function


	'--------------------------------------------------------------------------------------------------------------------
	<ExcelFunction(Description:="Получить сохранённые значения массива по его имени", Category:="ANik")>
	Public Function GetArr(<ExcelArgument(Description:="Имя массива")> arrName As String,
								  <ExcelArgument(Description:="Пропустить спереди")> Optional skip As Object = Nothing,
								  <ExcelArgument(Description:="Взять заданное количество")> Optional take As Object = Nothing)
		Dim oArr As Object
		If Not DictOfArrays.TryGetValue(arrName, oArr) Then
			Return Excel.XlCVError.xlErrValue
		End If

		Dim iSkip, iTake As Integer
		Dim errSkip, errTake As Boolean
		Dim srcArr = CType(oArr, Array)

		Try
			iSkip = CInt(skip)
		Catch
			errSkip = True
			iSkip = 0
		End Try

		Try
			iTake = CInt(take)
		Catch
			errTake = True
			iTake = srcArr.Length - iSkip
		End Try

		'если не заданы skip и take, массив вернётся как есть
		If errTake And errSkip Then

			'If TypeOf skip Is ExcelMissing And TypeOf take Is ExcelMissing Then
			Return oArr
		End If

		Try
			'Подготовка к созданию результирующего массива
			' поправить take, если он задан неверно
			If iTake > srcArr.Length - iSkip Then iTake = srcArr.Length - iSkip

			'определить кол-во размерностей и создать результирующий массив
			Dim T As Type
			Dim destArr As Object
			If srcArr.Rank = 1 Then
				T = srcArr(0).GetType()
				destArr = Array.CreateInstance(T, iTake)

			ElseIf srcArr.Rank = 2 Then
				T = srcArr(0, 0).GetType()
				destArr = Array.CreateInstance(T, 1, CType(iTake, Int16))

			Else
				Return oArr
			End If

			'заполнить результирующий массив
			Array.Copy(srcArr, iSkip, destArr, 0, iTake)
			Return destArr

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try
	End Function


	'--------------------------------------------------------------------------------------------------------------------
	<ExcelFunction(Description:="Транспонировать массив и сохранить его в словаре под заданным именем", Category:="ANik")
	>'-------------------------------------------------------------------------------------------------------------------
	Public Function SetTranspArr(<ExcelArgument(Description:="Имя сохраняемого массива")> arrName As String,
										  <ExcelArgument(Description:="Массив")> arr As Object,
										  <ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого")> Optional ReturnValue As Object = Nothing)
		Try
			Dim srcArr = CType(arr, Array)

			'определить кол-во размерностей и создать результирующий массив
			Dim destArr As Object
			Dim T As Type

			'если массив одномерный - сделать из него двумерный с одной строкой
			If srcArr.Rank = 1 Then
				T = srcArr(LBound(srcArr, 1)).GetType()

				'если массив текстовый, заполнять придётся вручную
				If TypeOf srcArr(LBound(srcArr, 1)) Is String Then
					ReDim destArr(0, srcArr.Length - 1)
					For i = 0 To srcArr.Length - 1
						destArr(0, i) = srcArr(i)
					Next
				Else
					Dim elSize = Len(srcArr(LBound(srcArr, 1)))
					destArr = Array.CreateInstance(T, 1, srcArr.Length)      'размерности зад-ся с нуля
					'Array.Copy(srcArr, 0, destArr, 0, srcArr.Length - 1)		'не работает, т.к. разное кол-во размерностей
					Buffer.BlockCopy(srcArr, 0, destArr, 0, (srcArr.Length) * elSize)
				End If

			ElseIf srcArr.Rank = 2 Then
				'транспонируемый массив - двумерный

				'тип первого элемента
				T = srcArr(LBound(srcArr, 1), LBound(srcArr, 2)).GetType()

				'если он в виде одной строки или столбца
				If srcArr.GetLength(0) = 1 Or srcArr.GetLength(1) = 1 Then
					'да, можно будет скопировать данные без цикла посредством Array.Copy()

					'если массив - одна строка
					If srcArr.GetLength(0) = 1 Then
						destArr = Array.CreateInstance(T, srcArr.Length, 1)     'размерности зад-ся с нуля
					Else
						'массив - один столбец
						destArr = Array.CreateInstance(T, 1, srcArr.Length)     'размерности зад-ся с нуля
					End If

					Array.Copy(srcArr, 0, destArr, 0, srcArr.Length)

				Else
					'массив - прямоугольная матрица, будет цикл

					'создать двумерный массив
					Dim srcRows = srcArr.GetLength(0)
					Dim srcCols = srcArr.GetLength(1)
					destArr = Array.CreateInstance(T, srcCols, srcRows)     'размерности зад-ся с нуля

					'проход по исходному массиву
					For rows = 0 To srcRows - 1
						For cols = 0 To srcCols - 1
							destArr(cols, rows) = srcArr(rows, cols)
						Next
					Next
				End If

			Else
				'это либо не массив, либо размерностей больше 2
				Return Excel.XlCVError.xlErrValue
			End If

			DictOfArrays(arrName) = destArr
			If TypeOf ReturnValue Is ExcelMissing Then Return destArr Else Return ReturnValue

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try
	End Function


	'--------------------------------------------------------------------------------------------------------------------
	<ExcelFunction(Description:="Получить строку массива", Category:="ANik")
	>'-------------------------------------------------------------------------------------------------------------------
	Function GetArrRow(<ExcelArgument(Description:="Имя переменной, где сохранён массив")> arrName As String,
							 <ExcelArgument(Description:="Номер строки, начиная с 1")> RowNumber As Integer)
		Dim oArr
		If Not DictOfArrays.TryGetValue(arrName, oArr) Then
			Return Excel.XlCVError.xlErrValue
		Else
			Dim Arr(,) As Object = CType(oArr, Array)
			'Arr = oArr

			'число столбцов массива
			Dim cols = Arr.GetLength(2) - 1

			Dim retArr(0 To cols)

			'скопировать нужный столбец
			For col = 0 To cols
				retArr(col) = Arr(RowNumber - 1, col)
			Next

			Return retArr
		End If

	End Function
	'--------------------------------------------------------------------------------------------------------------------

#Region "  SetArr 1..5 "

	Private Arr1
	<ExcelFunction(Description:="Сохранить значения массива в переменную", Category:="ANik")>
	Public Function SetArr1(<ExcelArgument(Description:="Комментарий назначения переменной")> Descr As String,
										<ExcelArgument(Description:="Массив")> VarValue() As Object,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого массива")> Optional ReturnValue As Object = Nothing)
		Arr1 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Arr2
	<ExcelFunction(Description:="Сохранить значения массива в переменную", Category:="ANik")>
	Public Function SetArr2(<ExcelArgument(Description:="Комментарий назначения переменной")> Descr As String,
										<ExcelArgument(Description:="Массив")> VarValue() As Object,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого массива")> Optional ReturnValue As Object = Nothing)
		Arr2 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Arr3
	<ExcelFunction(Description:="Сохранить значения массива в переменную", Category:="ANik")>
	Public Function SetArr3(<ExcelArgument(Description:="Комментарий назначения переменной")> Descr As String,
										<ExcelArgument(Description:="Массив")> VarValue() As Object,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого массива")> Optional ReturnValue As Object = Nothing)
		Arr3 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Arr4
	<ExcelFunction(Description:="Сохранить значения массива в переменную", Category:="ANik")>
	Public Function SetArr4(<ExcelArgument(Description:="Комментарий назначения переменной")> Descr As String,
										<ExcelArgument(Description:="Массив")> VarValue() As Object,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого массива")> Optional ReturnValue As Object = Nothing)
		Arr4 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

	Private Arr5
	<ExcelFunction(Description:="Сохранить значения массива в переменную", Category:="ANik")>
	Public Function SetArr5(<ExcelArgument(Description:="Комментарий назначения переменной")> Descr As String,
										<ExcelArgument(Description:="Массив")> VarValue() As Object,
										<ExcelArgument(Description:="Необязательное возвращаемое значение вместо сохраняемого массива")> Optional ReturnValue As Object = Nothing)
		Arr5 = VarValue
		If TypeOf ReturnValue Is ExcelMissing Then Return VarValue Else Return ReturnValue
	End Function

#End Region

#Region "  GetArr 1..5 "

	<ExcelFunction(Description:="Получить сохранённые значения массива", Category:="ANik")>
	Public Function GetArr1(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Пропустить спереди")> Optional Skip As Integer = 0,
										<ExcelArgument(Description:="Взять заданное количество")> Optional Take As Integer = 0)

		'если не заданы skip и take, массив вернётся как есть
		If Skip = -1 And Take = -1 Then
			Return Arr1
		End If

		Try
			Dim srcArr = CType(Arr1, Array)

			'Подготовка к созданию результирующего массива
			' вычислить take, если он не задан
			If Take = -1 Then
				Take = srcArr.Length - Skip
			Else
				Take = If(Take > (srcArr.Length - Skip), srcArr.Length - Skip, Take)
			End If

			'определить кол-во размерностей и создать результирующий массив
			Dim T As Type
			Dim destArr As Object
			If srcArr.Rank = 1 Then
				T = srcArr(0).GetType()
				destArr = Array.CreateInstance(T, Take)

			ElseIf srcArr.Rank = 2 Then
				T = srcArr(0, 0).GetType()
				destArr = Array.CreateInstance(T, 1, Take)

			Else
				Return Arr1
			End If

			'заполнить результирующий массив
			Array.Copy(srcArr, Skip, destArr, 0, Take)
			Return destArr

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try

	End Function

	<ExcelFunction(Description:="Получить сохранённые значения массива", Category:="ANik")>
	Public Function GetArr2(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Пропустить спереди")> Optional Skip As Integer = 0,
										<ExcelArgument(Description:="Взять заданное количество")> Optional Take As Integer = 0)

		'если не заданы skip и take, массив вернётся как есть
		If Skip = -1 And Take = -1 Then
			Return Arr2
		End If

		Try
			Dim srcArr = CType(Arr2, Array)

			'Подготовка к созданию результирующего массива
			' вычислить take, если он не задан
			If Take = -1 Then
				Take = srcArr.Length - Skip
			Else
				Take = If(Take > (srcArr.Length - Skip), srcArr.Length - Skip, Take)
			End If

			'определить кол-во размерностей и создать результирующий массив
			Dim T As Type
			Dim destArr As Object
			If srcArr.Rank = 1 Then
				T = srcArr(0).GetType()
				destArr = Array.CreateInstance(T, Take)

			ElseIf srcArr.Rank = 2 Then
				T = srcArr(0, 0).GetType()
				destArr = Array.CreateInstance(T, 1, Take)

			Else
				Return Arr2
			End If

			'заполнить результирующий массив
			Array.Copy(srcArr, Skip, destArr, 0, Take)
			Return destArr

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try

	End Function

	<ExcelFunction(Description:="Получить сохранённые значения массива", Category:="ANik")>
	Public Function GetArr3(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Пропустить спереди")> Optional Skip As Integer = 0,
										<ExcelArgument(Description:="Взять заданное количество")> Optional Take As Integer = 0)

		'если не заданы skip и take, массив вернётся как есть
		If Skip = -1 And Take = -1 Then
			Return Arr3
		End If

		Try
			Dim srcArr = CType(Arr3, Array)

			'Подготовка к созданию результирующего массива
			' вычислить take, если он не задан
			If Take = -1 Then
				Take = srcArr.Length - Skip
			Else
				Take = If(Take > (srcArr.Length - Skip), srcArr.Length - Skip, Take)
			End If

			'определить кол-во размерностей и создать результирующий массив
			Dim T As Type
			Dim destArr As Object
			If srcArr.Rank = 1 Then
				T = srcArr(0).GetType()
				destArr = Array.CreateInstance(T, Take)

			ElseIf srcArr.Rank = 2 Then
				T = srcArr(0, 0).GetType()
				destArr = Array.CreateInstance(T, 1, Take)

			Else
				Return Arr3
			End If

			'заполнить результирующий массив
			Array.Copy(srcArr, Skip, destArr, 0, Take)
			Return destArr

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try

	End Function

	<ExcelFunction(Description:="Получить сохранённые значения массива", Category:="ANik")>
	Public Function GetArr4(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Пропустить спереди")> Optional Skip As Integer = 0,
										<ExcelArgument(Description:="Взять заданное количество")> Optional Take As Integer = 0)

		'если не заданы skip и take, массив вернётся как есть
		If Skip = -1 And Take = -1 Then
			Return Arr4
		End If

		Try
			Dim srcArr = CType(Arr4, Array)

			'Подготовка к созданию результирующего массива
			' вычислить take, если он не задан
			If Take = -1 Then
				Take = srcArr.Length - Skip
			Else
				Take = If(Take > (srcArr.Length - Skip), srcArr.Length - Skip, Take)
			End If

			'определить кол-во размерностей и создать результирующий массив
			Dim T As Type
			Dim destArr As Object
			If srcArr.Rank = 1 Then
				T = srcArr(0).GetType()
				destArr = Array.CreateInstance(T, Take)

			ElseIf srcArr.Rank = 2 Then
				T = srcArr(0, 0).GetType()
				destArr = Array.CreateInstance(T, 1, Take)

			Else
				Return Arr4
			End If

			'заполнить результирующий массив
			Array.Copy(srcArr, Skip, destArr, 0, Take)
			Return destArr

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try

	End Function

	<ExcelFunction(Description:="Получить сохранённые значения массива", Category:="ANik")>
	Public Function GetArr5(<ExcelArgument(Description:="Комментарий назначения переменной")> Optional Descr As String = "",
										<ExcelArgument(Description:="Пропустить спереди")> Optional Skip As Integer = 0,
										<ExcelArgument(Description:="Взять заданное количество")> Optional Take As Integer = 0)

		'если не заданы skip и take, массив вернётся как есть
		If Skip = -1 And Take = -1 Then
			Return Arr5
		End If

		Try
			Dim srcArr = CType(Arr5, Array)

			'Подготовка к созданию результирующего массива
			' вычислить take, если он не задан
			If Take = -1 Then
				Take = srcArr.Length - Skip
			Else
				Take = If(Take > (srcArr.Length - Skip), srcArr.Length - Skip, Take)
			End If

			'определить кол-во размерностей и создать результирующий массив
			Dim T As Type
			Dim destArr As Object
			If srcArr.Rank = 1 Then
				T = srcArr(0).GetType()
				destArr = Array.CreateInstance(T, Take)

			ElseIf srcArr.Rank = 2 Then
				T = srcArr(0, 0).GetType()
				destArr = Array.CreateInstance(T, 1, Take)

			Else
				Return Arr5
			End If

			'заполнить результирующий массив
			Array.Copy(srcArr, Skip, destArr, 0, Take)
			Return destArr

		Catch
			Return Excel.XlCVError.xlErrValue
		End Try

	End Function

#End Region

#End Region



	<ExcelFunction(Description:="Вывод результата макро функции", Category:="ANik")>
	Function Macro(<ExcelArgument(Description:="Тело макро функции")> FormuLa, <ExcelArgument(Description:="Формула для вывода результата")> Optional Result = Nothing)

		If TypeOf Result Is ExcelMissing Then
			Return ExcelDnaUtil.Application.Caller.FormulaLocal()
		Else
			Return Result
		End If

	End Function


	<ExcelFunction(Description:="Сцепить два массива", Category:="ANik")>
	Function ConcatArr(<ExcelArgument(Description:="Массив 1")> Arr1(), <ExcelArgument(Description:="Массив 2")> Arr2())

		Return Arr1.Concat(Arr2).ToArray()

	End Function


	<ExcelFunction(Description:="My first .NET function", Category:="ANik")>
	Public Function HelloR(<ExcelArgument(Description:="Имя кого поприветствовать")> Name As String) As String

		Return "Hello " & Name

	End Function

End Module

