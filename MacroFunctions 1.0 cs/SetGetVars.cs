using System;

using ExcelDna.Integration;
using Microsoft.Office.Interop;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using static Microsoft.VisualBasic.Interaction;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;

using System.Security;

namespace MacroFunctions
{
    public static class SetGetVars
    {
        [DllImport("Kernel32.dll"), SuppressUnmanagedCodeSecurity]
        public static extern int GetCurrentProcessorNumber();

        private static int[,] intVar;
        private static double[,] doubVar;
        private static string[,] strVar;
        private static object[,] objVar;

        static SetGetVars()
        {
            int procCount = Environment.ProcessorCount;

            //создать массивы для хранения переменных для каждого ядра
            intVar = new int[16, procCount];
            doubVar = new double[16, procCount];
            strVar = new string[16, procCount];
            objVar = new object[16, procCount];
        }

        #region Call UDF from XLL надстройки 

        public static Dictionary<string, object> dictOfWSheetObjects = new Dictionary<string, object>();
        const string _shName = "[TblWork_1.2.xlam]";

        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Вызвать указанную UDF, определённую в той же книге, методом Run. Имя книги берётся из App.Caller", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object RunFunc_pa(string FuncName, params object[] Arr)
        {   //               ``````````
            var app = (Application)ExcelDnaUtil.Application;
            var callerCell = (Range)app.Caller;
            string callerCellAddr = callerCell.Address[true, true, XlReferenceStyle.xlA1, true];
            string wbName = callerCellAddr.Substring(1, callerCellAddr.IndexOf("]") - 1).Replace("[", "").Replace("'", "");

            string wbFullName = "'" + wbName + "'!" + FuncName;
            return app.Run(wbFullName, Arr);
        }


        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Вызвать указанную UDF, определённую в той же книге, методом Run. Имя книги берётся из App.Caller", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object RunFunc(string FuncName, object par1 = null, object par2 = null, object par3 = null, object par4 = null, object par5 = null)
        {   //               ```````
            var app = (Application)ExcelDnaUtil.Application;
            var callerCell = (Range)app.Caller;
            string callerCellAddr = callerCell.Address[true, true, XlReferenceStyle.xlA1, true];
            string wbName = callerCellAddr.Substring(1, callerCellAddr.IndexOf("]") - 1).Replace("[", "").Replace("'", "");

            string wbFullName = "'" + wbName + "'!" + FuncName;

            if (par1 is ExcelMissing)
            {
                return app.Run(wbFullName);
            }
            else if (par2 is ExcelMissing)
            {
                return app.Run(wbFullName, par1);
            }
            else if (par3 is ExcelMissing)
            {
                return app.Run(wbFullName, par1, par2);
            }
            else if (par4 is ExcelMissing)
            {
                return app.Run(wbFullName, par1, par2, par3);
            }
            else if (par5 is ExcelMissing)
            {
                return app.Run(wbFullName, par1, par2, par3, par4);
            }
            else
            {
                return app.Run(wbFullName, par1, par2, par3, par4, par5);
            }
        }


        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Вызвать указанную UDF либо из надстройки [TblWork_1.2.xlam] (тогда shName=''), либо из модуля рабочего листа (тогда shName='[имяКниги.расш]ИмяЛиста')", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object CallBN(string shName, string FuncName, object par1 = null, object par2 = null, object par3 = null, object par4 = null, object par5 = null)
        {
            var app = (Application)ExcelDnaUtil.Application;
            object wsObject;
            string wbName;

            //•определить, откуда требуется вызвать функцию: из надстройки или с листа, на котором она расположена
            // если имя листа задано
            if (shName != "")
            {
                //вызвать функцию с листа, откуда пришёл вызов
                //--------------------------------------------

                //проверить наличие в словаре ссылки на лист, с которого пришёл вызов
                if (dictOfWSheetObjects.TryGetValue(shName, out wsObject) == false)
                {
                    //в словаре ссылки на кнмгу/лист нет, добавить её
                    //-----------------------------------------------

                    //выделить из параметра shName название книги и листа
                    wbName = shName.Substring(2, shName.IndexOf("]") - 2);
                    object wsName = shName.Substring(shName.IndexOf("]") + 1, 100);
                    wsObject = (Worksheet)app.Workbooks[wbName].Worksheets[wsName];

                    //добавить в словарь лист, откуда пришёл вызов
                    dictOfWSheetObjects[shName] = wsObject;
                }
                else
                {
                    //в словаре ссылка на лист есть, получена в wsObject
                }
            }
            else
            {
                //имя листа не задано, вызвать функции из надстройки
                //--------------------------------------------------

                //проверить наличие ключа c именем надстройки в словаре
                if (dictOfWSheetObjects.TryGetValue(_shName, out wsObject) == false)
                {
                    //в словаре ссылки на надстройку нет, добавить её
                    //-----------------------------------------------

                    //выделить из параметра shName название книги и листа
                    wbName = _shName.Substring(2, _shName.IndexOf("]") - 2);    //wsName = Mid(_shName, InStr(_shName, "]") + 1, 100) 
                    wsObject = app.Workbooks[wbName];

                    //добавить в словарь лист, откуда пришёл вызов
                    dictOfWSheetObjects[_shName] = wsObject;
                }
                else
                {
                    //в словаре ссылка на надстройку есть, получена в wsObject}
                }
            }
            if (par1 is ExcelMissing)
            {
                return CallByName(wsObject, FuncName, CallType.Method);
            }
            else if (par2 is ExcelMissing)
            {
                return CallByName(wsObject, FuncName, CallType.Method, par1);
            }
            else if (par3 is ExcelMissing)
            {
                return CallByName(wsObject, FuncName, CallType.Method, par1, par2);
            }
            else if (par4 is ExcelMissing)
            {
                return CallByName(wsObject, FuncName, CallType.Method, par1, par2, par3);
            }
            else if (par5 is ExcelMissing)
            {
                return CallByName(wsObject, FuncName, CallType.Method, par1, par2, par3, par4);
            }
            else
            {
                return CallByName(wsObject, FuncName, CallType.Method, par1, par2, par3, par4, par5);
            }

        }

        #endregion


        #region SetVar / GetVar / ClearVarDict

        private static Dictionary<string, object> dictOfVars = new Dictionary<string, object>();

        [ExcelFunction(Description = "Задать значение переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetVar([ExcelArgument(Description = "Имя переменной")] string varName,
                             [ExcelArgument(Description = "Значение")] object varValue,
                             [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            dictOfVars[varName] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        [ExcelFunction(Description = "Получить значение переменной по имени", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetVar([ExcelArgument(Description = "Имя переменной")] string varName)
        {
            if (!dictOfVars.TryGetValue(varName, out object varValue))
            {
                return XlCVError.xlErrValue;
            }
            else
            {
                return varValue;
            }
        }

        [ExcelFunction(Description = "Удалить все именные переменные", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object ClearVarDict(object returnValue = null)
        {
            dictOfVars.Clear();
            dictOfArrays.Clear();
            intVar.Initialize();
            doubVar.Initialize();
            strVar.Initialize();
            objVar.Initialize();
            if (returnValue == null || returnValue is ExcelMissing)
                return 0;
            else
                return returnValue;
        }

        #endregion


        #region SetInt / GetIntN

        [ExcelFunction(Description = "Задать значение одной из 16 целочисленных переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetInt(
                [ExcelArgument(Description = "Номер переменной от 0..15")]         /**/ int varNumber,
                [ExcelArgument(Description = "Комментарий назначения переменной")] /**/ object descr = null,
                [ExcelArgument(Description = "Значение")]                          /**/ int varValue = 0,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            intVar[varNumber, 0] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить значение одной из 16 целочисленных переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetInt(
                [ExcelArgument(Description = "Номер переменной от 0..15")] /**/    int varNumber = 0,
                [ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return intVar[varNumber, 0];
        }


        [ExcelFunction(Description = "Задать значение одной из 16 целочисленных переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetTmpInt(
                [ExcelArgument(Description = "Номер переменной от 0..15")]         /**/ int varNumber,
                [ExcelArgument(Description = "Комментарий назначения переменной")] /**/ object descr = null,
                [ExcelArgument(Description = "Значение")]                          /**/ int varValue = 0,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            intVar[varNumber, GetCurrentProcessorNumber()] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить значение одной из 16 целочисленных переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetTmpInt(
                [ExcelArgument(Description = "Номер переменной от 0..15")] /**/    int varNumber = 0,
                [ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return intVar[varNumber, GetCurrentProcessorNumber()];
        }

        #endregion


        #region SetDoubN / GetDoubN

        [ExcelFunction(Description = "Задать значение одной из 16 дробных (вещественных) переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetDoub(
                [ExcelArgument(Description = "Комментарий назначения переменной")] /**/ object descr = null,
                [ExcelArgument(Description = "Номер переменной от 0..15")]         /**/ int varNumber = 0,
                [ExcelArgument(Description = "Значение")]                          /**/ double varValue = 0,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            doubVar[varNumber, GetCurrentProcessorNumber()] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        
        [ExcelFunction(Description = "Получить значение одной из 16 дробных (вещественных) переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetDoub([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                                      [ExcelArgument(Description = "Номер переменной от 0..15")] /**/    int varNumber = 0)
        {
            return doubVar[varNumber, GetCurrentProcessorNumber()];
        }

        #endregion


        #region SetStrN / GetStrN

        [ExcelFunction(Description = "Задать значение одной из 16 текстовых переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetStr(
                [ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                [ExcelArgument(Description = "Номер переменной от 0..15")]    /**/ int varNumber = 0,
                [ExcelArgument(Description = "Значение")]                     /**/ string varValue = "",
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            strVar[varNumber, GetCurrentProcessorNumber()] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        [ExcelFunction(Description = "Получить значение одной из 16 текстовых переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetStr([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                                     [ExcelArgument(Description = "Номер переменной от 0..15")] /**/    int varNumber = 0)
        {
            return strVar[varNumber, GetCurrentProcessorNumber()];
        }

        #endregion


        #region SetArr/SetArrN, GetArr/GetArrN, GetArrRow

        private static Dictionary<string, object> dictOfArrays = new Dictionary<string, object>();
        private static Type elementType;

        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Сохранить имя и значения массива в словаре", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr(
                [ExcelArgument(Description = "Имя сохраняемого массива")] string arrName,
                [ExcelArgument(Description = "Массив")] object Arr,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            dictOfArrays[arrName] = Arr;
            if (returnValue == null || returnValue is ExcelMissing)
                return Arr;
            else
                return returnValue;
        }


        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Транспонировать массив и сохранить его в словаре под заданным именем", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetTranspArr(
                [ExcelArgument(Description = "Имя сохраняемого массива")] string arrName,
                [ExcelArgument(Description = "Массив")] object array,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            try
            {
                //определить кол-во размерностей и создать результирующий массив
                //Type T;

                if (!array.GetType().IsArray) return array;

                Array srcArr = (Array)array;
                Array destArr;

                //если массив одномерный сделать из него двумерный с одной строкой
                if (srcArr.Rank == 1)
                {
                    elementType = srcArr.GetType().GetElementType();

                    //если массив текстовый, заполнять придётся вручную
                    if (array is string[] txtSrcArr)
                    {
                        string[,] txtDstArr = new string[1, srcArr.Length];

                        for (int i = 0; i <= srcArr.Length - 1; i++)
                        {
                            txtDstArr[0, i] = txtSrcArr[i];
                        }
                        destArr = txtDstArr;
                    }
                    else
                    {// массив не текстовый, заполнить значениями можно исп. Buffer.BlockCopy

                        int elSize = Marshal.SizeOf(elementType);

                        destArr = Array.CreateInstance(elementType, 1, srcArr.Length);        //размерности зад-ся с нуля
                        Buffer.BlockCopy(srcArr, 0, destArr, 0, (srcArr.Length) * elSize);

                        //Array.Copy(srcArr, 0, destArr, 0, srcArr.Length - 1)		//не работает, т.к.разное кол-во размерностей
                    }
                }
                else if (srcArr.Rank == 2)
                {// транспонируемый массив - двумерный

                    //тип элементов массива
                    elementType = srcArr.GetType().GetElementType();

                    //если он в виде одной строки или столбца
                    if (srcArr.GetLength(0) == 1 | srcArr.GetLength(1) == 1)
                    {
                        //можно будет скопировать данные без цикла посредством Array.Copy()

                        //если массив - одна строка
                        if (srcArr.GetLength(0) == 1)
                        {
                            destArr = Array.CreateInstance(elementType, srcArr.Length, 1);    //размерности зад-ся с нуля
                        }
                        else
                        {// массив - один столбец
                            destArr = Array.CreateInstance(elementType, 1, srcArr.Length);    //размерности зад-ся с нуля
                        }
                        Array.Copy(srcArr, 0, destArr, 0, srcArr.Length);
                    }
                    else
                    {// массив - прямоугольная матрица, будет цикл для преобразования

                        //создать двумерный массив
                        int srcRows = srcArr.GetLength(0);
                        int srcCols = srcArr.GetLength(1);
                        dynamic tmdDestArr = Array.CreateInstance(elementType, srcCols, srcRows);      //размерности зад-ся с нуля
                        dynamic tmpSrcArr = array;

                        //заполнить созданный массив транспонированными данными исходного
                        for (int row = 0; row <= srcRows - 1; row++)
                        {
                            for (int col = 0; col <= srcCols - 1; col++)
                            {
                                tmdDestArr[col, row] = tmpSrcArr[row, col];
                            }
                        }
                        destArr = tmdDestArr;
                    }
                }
                else
                {// это либо не массив, либо размерностей больше 2
                    return XlCVError.xlErrValue;
                }

                dictOfArrays[arrName] = destArr;
                if (returnValue == null || returnValue is ExcelMissing)
                    return destArr;
                else
                    return returnValue;
            }
            catch
            {
                return XlCVError.xlErrValue;
            }
        }


        /*------------------------------------------------------------------------------------------------------------*/
        private static object GetArrayPart(object oArr, object skip = null, object take = null)
        {
            int iSkip, iTake;
            bool errSkip = false, errTake = false;

            if (!oArr.GetType().IsArray) return oArr;

            var srcArr = (Array)oArr;

            if (skip == null)
            {
                if (take == null)
                {
                    return oArr;
                }
                iSkip = 0;
            }

            try
            {
                iSkip = Convert.ToInt32(skip);
            }
            catch
            {
                errSkip = true;
                iSkip = 0;
            }

            try
            {
                if (take == null)
                {
                    iTake = srcArr.Length - iSkip;
                }
                iTake = Convert.ToInt32(take);
            }
            catch
            {
                errTake = true;
                iTake = srcArr.Length - iSkip;
            }

            //если skip или take невозможно преобр. в int - вернуть ошибку
            if (errSkip || errTake)
            {
                return XlCVError.xlErrValue;
            }

            try
            {
                //проверка на корректность iTake
                if (iTake > srcArr.Length - iSkip || iTake < 1)
                {
                    iTake = srcArr.Length - iSkip;  //вернусть весь массив
                }

                //создать результирующий массив
                Array destArr;
                elementType = oArr.GetType().GetElementType();

                if (srcArr.Rank == 1)
                {
                    destArr = Array.CreateInstance(elementType, iTake);
                }
                else if (srcArr.Rank == 2)
                {
                    destArr = Array.CreateInstance(elementType, 1, iTake);
                }
                else
                {
                    return oArr;    //если вдруг размерностей больше 2х
                }

                Array.Copy(srcArr, iSkip, destArr, 0, iTake);
                return destArr;
            }
            catch
            {
                return XlCVError.xlErrValue;
            }
        }


        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Получить сохранённые значения массива по его имени", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr(
                      [ExcelArgument(Description = "Имя массива")] string arrName,
                      [ExcelArgument(Description = "Пропустить спереди")] object skip = null,
                      [ExcelArgument(Description = "Взять заданное количество")] object take = null)
        {
            if (!dictOfArrays.TryGetValue(arrName, out object oArr))
            {
                return XlCVError.xlErrValue;
            }

            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return oArr;
            }

            if (!oArr.GetType().IsArray) return oArr;

            return GetArrayPart(oArr, skip, take);
        }

        /*------------------------------------------------------------------------------------------------------------*/
        [ExcelFunction(Description = "Получить строку массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArrRow([ExcelArgument(Description = "Имя переменной, где сохранён массив")] string arrName,
                                        [ExcelArgument(Description = "Номер строки, начиная с 1")] int rowNumber)
        {
            //извлечь массив из словаря
            if (!dictOfArrays.TryGetValue(arrName, out object oArr))
            {
                return XlCVError.xlErrValue;
            }

            if (!oArr.GetType().IsArray) return oArr;

            dynamic srcArr = oArr;

            //если массив одномерный - вернуть элемент по указанному индексу
            if (srcArr.Rank == 1)
                return srcArr[rowNumber - 1];   //т.к. решили строки номеровать с 1

            //создать одномерный массив под извлекаемую строку
            elementType = oArr.GetType().GetElementType();
            int cols = srcArr.GetLength(1);
            dynamic arrRow = Array.CreateInstance(elementType, cols);

            for (int col = 0; col < cols; col++)
            {
                arrRow[col] = srcArr[rowNumber - 1, col];
            }
            return arrRow;
        }


        #region SetArrN / GetArrN

        [ExcelFunction(Description = "Сохранить значения массива в одну из 16 переменных", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr1(
                [ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                [ExcelArgument(Description = "Номер переменной от 0..15")]    /**/ int varNumber = 0,
                [ExcelArgument(Description = "Массив")]                       /**/ object varValue = null,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            objVar[varNumber, GetCurrentProcessorNumber()] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить сохранённые значения массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr1(
                [ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                [ExcelArgument(Description = "Номер переменной от 0..15")]    /**/ int varNumber = 0,
                [ExcelArgument(Description = "Пропустить спереди")]           /**/ object skip = null,
                [ExcelArgument(Description = "Взять заданное количество")]    /**/ object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return objVar[varNumber, GetCurrentProcessorNumber()];
            }
            int currProcNum = GetCurrentProcessorNumber();
            if (!objVar[varNumber, currProcNum].GetType().IsArray) return objVar[varNumber, currProcNum];

            return GetArrayPart(objVar[varNumber, currProcNum], skip, take);
        }

        #endregion

        #endregion


        [ExcelFunction(Description = "Вывод результата макро функции", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object Macro([ExcelArgument(Description = "Тело макро функции")] object FormuLa,
                            [ExcelArgument(Description = "Формула для вывода результата")] object Result = null)
        {
            if (Result is ExcelMissing)
            {
                return ((Application)ExcelDnaUtil.Application).Caller.FormulaLocal();
            }
            else
            {
                return Result;
            }
        }


        [ExcelFunction(Description = "Сцепить два массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object ConcatArr([ExcelArgument(Description = "Массив 1")] object[] arr1, [ExcelArgument(Description = "Массив 2")] object[] arr2)
        {
            return arr1.Concat(arr2).ToArray();
        }

        [ExcelFunction(Description = "My first .NET function", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static string HelloR([ExcelArgument(Description = "Имя кого поприветствовать")] string Name)
        {
            return "Hello " + Name;
        }
    }
}
