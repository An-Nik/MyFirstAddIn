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

namespace MacroFunctions
{
    public static class SetGetVars
    {
        /*static SetGetVars()
        {
            int procCount = Environment.ProcessorCount;

            //
            int[,] intArr = new int[8, procCount];
            double[,] doubArr = new double[8, procCount];
            string[,] strArr = new string[8, procCount];
            object[,] objArr = new object[8, procCount];
        }*/

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
            if (returnValue == null || returnValue is ExcelMissing)
                return 0;
            else
                return returnValue;
        }

        #endregion


        #region SetIntN / GetIntN

        private static long int1;
        [ExcelFunction(Description = "Задать значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetInt1(
                [ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                [ExcelArgument(Description = "Значение")] int varValue = 0,
                [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            int1 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static long int2;
        [ExcelFunction(Description = "Задать значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetInt2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] int varValue = 0,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            int2 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static long int3;
        [ExcelFunction(Description = "Задать значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetInt3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] int varValue = 0,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            int3 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static long int4;
        [ExcelFunction(Description = "Задать значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetInt4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] int varValue = 0,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            int4 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static long int5;
        [ExcelFunction(Description = "Задать значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetInt5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] int varValue = 0,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            int5 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить сохранённое значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetInt1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return int1;
        }

        [ExcelFunction(Description = "Получить сохранённое значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetInt2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return int2;
        }

        [ExcelFunction(Description = "Получить сохранённое значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetInt3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return int3;
        }

        [ExcelFunction(Description = "Получить сохранённое значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetInt4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return int4;
        }

        [ExcelFunction(Description = "Получить сохранённое значение целочисленной переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetInt5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return int5;
        }

        #endregion


        #region SetDoubN / GetDoubN

        private static double doub1;
        [ExcelFunction(Description = "Задать значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetDoub1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                               [ExcelArgument(Description = "Значение")] double varValue = 0,
                               [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            doub1 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static double doub2;
        [ExcelFunction(Description = "Задать значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetDoub2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                               [ExcelArgument(Description = "Значение")] double varValue = 0,
                               [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            doub2 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static double doub3;
        [ExcelFunction(Description = "Задать значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetDoub3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                               [ExcelArgument(Description = "Значение")] double varValue = 0,
                               [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            doub3 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static double doub4;
        [ExcelFunction(Description = "Задать значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetDoub4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                               [ExcelArgument(Description = "Значение")] double varValue = 0,
                               [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            doub4 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static double doub5;
        [ExcelFunction(Description = "Задать значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetDoub5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                               [ExcelArgument(Description = "Значение")] double varValue = 0,
                               [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            doub5 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        [ExcelFunction(Description = "Получить сохранённое значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetDoub1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return doub1;
        }

        [ExcelFunction(Description = "Получить сохранённое значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetDoub2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return doub2;
        }

        [ExcelFunction(Description = "Получить сохранённое значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetDoub3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return doub3;
        }

        [ExcelFunction(Description = "Получить сохранённое значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetDoub4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return doub4;
        }

        [ExcelFunction(Description = "Получить сохранённое значение дробной (вещественной) переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetDoub5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return doub5;
        }

        #endregion


        #region SetStrN / GetStrN

        private static string str1;
        [ExcelFunction(Description = "Задать значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetStr1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] string varValue = "",
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            str1 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static string str2;
        [ExcelFunction(Description = "Задать значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetStr2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] string varValue = "",
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            str2 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static string str3;
        [ExcelFunction(Description = "Задать значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetStr3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] string varValue = "",
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            str3 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static string str4;
        [ExcelFunction(Description = "Задать значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetStr4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] string varValue = "",
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            str4 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static string str5;
        [ExcelFunction(Description = "Задать значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetStr5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Значение")] string varValue = "",
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            str5 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить сохранённое значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetStr1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return str1;
        }

        [ExcelFunction(Description = "Получить сохранённое значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetStr2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return str2;
        }

        [ExcelFunction(Description = "Получить сохранённое значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetStr3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return str3;
        }

        [ExcelFunction(Description = "Получить сохранённое значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetStr4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return str4;
        }

        [ExcelFunction(Description = "Получить сохранённое значение текстовой переменной", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetStr5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null)
        {
            return str5;
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

        #region SetArr 1..5

        private static object arr1;
        [ExcelFunction(Description = "Сохранить значения массива в переменную", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Массив")] object varValue = null,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            arr1 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static object arr2;
        [ExcelFunction(Description = "Сохранить значения массива в переменную", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Массив")] object varValue = null,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            arr2 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static object arr3;
        [ExcelFunction(Description = "Сохранить значения массива в переменную", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Массив")] object varValue = null,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            arr3 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static object arr4;
        [ExcelFunction(Description = "Сохранить значения массива в переменную", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Массив")] object varValue = null,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            arr4 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        private static object arr5;
        [ExcelFunction(Description = "Сохранить значения массива в переменную", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object SetArr5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Массив")] object varValue = null,
                              [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            arr5 = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        #endregion

        #region GetArr 1..5

        [ExcelFunction(Description = "Получить сохранённые значения массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr1([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Пропустить спереди")] object skip = null,
                              [ExcelArgument(Description = "Взять заданное количество")] object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return arr1;
            }

            if (!arr1.GetType().IsArray) return arr1;

            return GetArrayPart(arr1, skip, take);
        }

        [ExcelFunction(Description = "Получить сохранённые значения массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr2([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Пропустить спереди")] object skip = null,
                              [ExcelArgument(Description = "Взять заданное количество")] object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return arr2;
            }

            if (!arr2.GetType().IsArray) return arr2;

            return GetArrayPart(arr2, skip, take);
        }

        [ExcelFunction(Description = "Получить сохранённые значения массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr3([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Пропустить спереди")] object skip = null,
                              [ExcelArgument(Description = "Взять заданное количество")] object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return arr3;
            }

            if (!arr3.GetType().IsArray) return arr3;

            return GetArrayPart(arr3, skip, take);
        }

        [ExcelFunction(Description = "Получить сохранённые значения массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr4([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Пропустить спереди")] object skip = null,
                              [ExcelArgument(Description = "Взять заданное количество")] object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return arr4;
            }

            if (!arr4.GetType().IsArray) return arr4;

            return GetArrayPart(arr4, skip, take);
        }

        [ExcelFunction(Description = "Получить сохранённые значения массива", Category = "ANik")]
        /*------------------------------------------------------------------------------------------------------------*/
        public static object GetArr5([ExcelArgument(Description = "Комментарий назначения переменной")] object descr = null,
                              [ExcelArgument(Description = "Пропустить спереди")] object skip = null,
                              [ExcelArgument(Description = "Взять заданное количество")] object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return arr5;
            }

            if (!arr5.GetType().IsArray) return arr5;

            return GetArrayPart(arr5, skip, take);
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