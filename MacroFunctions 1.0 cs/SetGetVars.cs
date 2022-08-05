﻿using System;
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
        //    [DllImport("Kernel32.dll"), SuppressUnmanagedCodeSecurity]
        //    public static extern int GetCurrentProcessorNumber();

        private static object[] objVar = new object[16];
        private static object[] objArr = new object[16];

        static SetGetVars()
        {
            //int procCount = Environment.ProcessorCount;

            //создать массивы для хранения переменных для каждого ядра
            //objVar = new object[16];
            //objArr = new object[16];
        }

        #region Call UDF from XLL надстройки 

        public static Dictionary<string, object> dictOfWSheetObjects = new Dictionary<string, object>();
        const string _shName = "[TblWork_1.2.xlam]";

        [ExcelFunction(Description = "Вызвать указанную UDF, определённую в той же книге, методом Run. Имя книги берётся из App.Caller", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object RunFunc_pa(string FuncName, params object[] Arr)
        //--------------------------------------------------------------------------------------------------------------
        {
            var app = (Application)ExcelDnaUtil.Application;
            var callerCell = (Range)app.Caller;
            string callerCellAddr = callerCell.Address[true, true, XlReferenceStyle.xlA1, true];
            string wbName = callerCellAddr.Substring(1, callerCellAddr.IndexOf("]") - 1).Replace("[", "").Replace("'", "");

            string wbFullName = "'" + wbName + "'!" + FuncName;
            return app.Run(wbFullName, Arr);
        }


        [ExcelFunction(Description = "Вызвать указанную UDF, определённую в той же книге, методом Run. Имя книги берётся из App.Caller", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object RunFunc(string FuncName, object par1 = null, object par2 = null, object par3 = null, object par4 = null, object par5 = null)
        //--------------------------------------------------------------------------------------------------------------
        {
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


        [ExcelFunction(Description = "Вызвать указанную UDF либо из надстройки [TblWork_1.2.xlam] (тогда shName=''), либо из модуля рабочего листа (тогда shName='[имяКниги.расш]ИмяЛиста')", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object CallBN(string shName, string FuncName, object par1 = null, object par2 = null, object par3 = null, object par4 = null, object par5 = null)
        //--------------------------------------------------------------------------------------------------------------
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


        #region NameSet, NameGet, NameDel

        private static Dictionary<string, object> dictOfVars = new Dictionary<string, object>();

        [ExcelFunction(Description = "Задать значение переменной по имени", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object NameSet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Имя переменной")] string varName,
            [ExcelArgument(Description = "Значение")]  /**/ object varValue,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            dictOfVars[varName] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }

        [ExcelFunction(Description = "Получить значение переменной по имени", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object NameGet([ExcelArgument(Description = "Имя переменной")] string varName)
        //--------------------------------------------------------------------------------------------------------------
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

        [ExcelFunction(Description = "Удалить переменную из словаря/очистить все переменные", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object NameDel(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Необязательное возвращаемое значение")]                           /**/ object returnValue = null,
            [ExcelArgument(Description = "Масив удаляемых имён переменной. Если не задано - очистит все имена")] params object[] names)
        {
            if (names[0] == null || names[0] is ExcelMissing)
                dictOfVars.Clear();
            else
            {
                foreach (var name in names)
                {
                    dictOfVars.Remove((string)name);
                }
            }
            if (returnValue == null || returnValue is ExcelMissing)
                return 0;
            else
                return returnValue;
        }

        #endregion


        #region VarSet, VarGet

        [ExcelFunction(Description = "Задать значение одной из 16 переменных", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object VarSet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер переменной от 0..15")]  /**/ int varNumber,
            [ExcelArgument(Description = "Значение")]                   /**/ object varValue,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого")] object returnValue = null)
        {
            objVar[varNumber] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить значение одной из 16 переменных", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object VarGet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер переменной от 0..15")]  /**/ int varNumber,
            [ExcelArgument(Description = "Название переменной/комментарий")] object descr = null)
        {
            return objVar[varNumber];
        }

        #endregion


        #region ArrayInit, ArraySet, ArrayGet, ArrayItemSet, ArrayItemGet, ArrayConcat, ArrayGetAs2d

        private static Type elementType;

        [ExcelFunction(Description = "Создать пустой одномерный массив", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrayInit(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] int varNumber,
            [ExcelArgument(Description = "Количество элементов")] /**/int count,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            objArr[varNumber] = new object[count];
            if (returnValue == null || returnValue is ExcelMissing)
                return objArr[varNumber];
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Сохранить массив в одну из 16 переменных", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArraySet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] int varNumber,
            [ExcelArgument(Description = "Массив")]              /**/ object varValue,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            objArr[varNumber] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить массив или часть массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrayGet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] /**/ int varNumber,
            [ExcelArgument(Description = "Пропустить спереди")]       /**/ object skip = null,
            [ExcelArgument(Description = "Извлекаемое количество")]   /**/ object take = null)
        {
            //если skip и take не заданы - вернуть полный массив
            if (skip is ExcelMissing) skip = null;
            if (take is ExcelMissing) take = null;

            if (skip == null && take == null)
            {
                return objArr[varNumber];
            }
            if (!objArr[varNumber].GetType().IsArray) return objArr[varNumber];

            return GetArrayPart(objArr[varNumber], skip, take);
        }


        [ExcelFunction(Description = "Задать значение элементу одномерного массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrayItemSet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")]        /**/ int varNumber,
            [ExcelArgument(Description = "Номер элемента массива (начиная с 1)")] int index,
            [ExcelArgument(Description = "Значение")]                        /**/ object varValue,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            //извлечь указанный массив
            dynamic array = objArr[varNumber];
            //object[] array = (object[])objArr[varNumber];
            array[index - 1] = varValue;

            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить значение элемента одномерного массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrayItemGet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")]         /**/ int varNumber,
            [ExcelArgument(Description = "Индекс элемента массива (начиная с 1)")] int index)
        {
            //извлечь нужный массив
            dynamic array = objArr[varNumber];
            //object[] array = (object[])objArr[varNumber];

            return array[index - 1];
        }


        [ExcelFunction(Description = "Сцепить два массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrayConcat(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Сцепляемый массив 1")] object[] arr1,
            [ExcelArgument(Description = "Сцепляемый массив 2")] object[] arr2)
        {
            //return arr1.Concat(arr2).ToArray();
            object[] resultArr = new object[arr1.Length + arr2.Length];
            Array.Copy(arr1, resultArr, arr1.Length);
            Array.Copy(arr2, 0, resultArr, arr1.Length, arr2.Length);
            return resultArr;
        }


        //--------------------------------------------------------------------------------------------------------------
        private static object GetArrayPart(object arraySrc, object skip = null, object take = null)
        //--------------------------------------------------------------------------------------------------------------
        {
            int iSkip, iTake;

            if (!arraySrc.GetType().IsArray) return arraySrc;

            var srcArr = (Array)arraySrc;

            try
            {
                //проверка корректности skip
                if (skip == null)
                {
                    if (take == null)
                    {
                        return arraySrc;
                    }
                    iSkip = 0;
                }
                else
                {
                    iSkip = Convert.ToInt32(skip);
                }

                //проверка корректности take
                if (take == null)
                {
                    iTake = srcArr.Length - iSkip;
                }
                else
                {
                    iTake = Convert.ToInt32(take);
                }

                //проверка корректности извлекаемого количества
                if (iTake > srcArr.Length - iSkip || iTake < 1)
                {
                    iTake = srcArr.Length - iSkip;
                }

                //создать результирующий массив
                Array destArr;
                elementType = arraySrc.GetType().GetElementType();

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
                    return XlCVError.xlErrValue;    //если вдруг размерностей больше 2х
                }

                Array.Copy(srcArr, iSkip, destArr, 0, iTake);
                return destArr;
            }
            catch
            {
                return XlCVError.xlErrValue;
            }
        }

        /*
        [ExcelFunction(Description = "Скопировать данные из одного массива в другой с заданной позиции", Category = "ANik")]
        //------------------------------------------------------------------------------------------------------------
        public static void ArrayCopy(
        //------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива источника (от 0..15)")] int srcNumber, 
            [ExcelArgument(Description = "Индекс начала данных (с 0)")]    /** / int srcSkip, 
            [ExcelArgument(Description = "Номер массива приёмника(от 0..15) ")] int destNumber,
            [ExcelArgument(Description = "Индекс начала данных (с 0)")]    /** / int destSkip,
            [ExcelArgument(Description = "Количество копируемых элементов")]/** /int take)
        {
            var dstArr = (Array)objArr[destNumber];
            var srcArr = (Array)objArr[srcNumber];

            try
            {
                //проверка корректности извлекаемого количества
                if (take > srcArr.Length - srcSkip || take < 1)
                {
                    take = srcArr.Length - srcSkip;
                    /* 
                    if (take > dstArr.Length - destSkip)
                    {
                        take = dstArr.Length - destSkip;
                    }* /
                }

                Array.Copy(srcArr, srcSkip, dstArr, destSkip, take);
            }
            catch
            {
                //return XlCVError.xlErrValue;
            }
        }*/

        #endregion


        #region ArrInit, ArrSet, ArrGet, ArrItemSet, ArrItemGet, ArrGetRow, ArrSetTransp, ArrGetAs1d

        [ExcelFunction(Description = "Создать пустой двумерный массив", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrInit(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] int varNumber,
            [ExcelArgument(Description = "Количество строк")]    /**/ int rows,
            [ExcelArgument(Description = "Количество столбцов")] /**/ int cols,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            objArr[varNumber] = new object[rows, cols];
            if (returnValue == null || returnValue is ExcelMissing)
                return objArr[varNumber];
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Сохранить массив в одну из 16 переменных", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrSet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] int varNumber,
            [ExcelArgument(Description = "Массив")]              /**/ object varValue,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            objArr[varNumber] = varValue;
            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить массив", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrGet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] /**/ int varNumber)
        {
            return objArr[varNumber];
        }


        [ExcelFunction(Description = "Задать значение элементу двумерного массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrItemSet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] /**/ int varNumber,
            [ExcelArgument(Description = "Строка (начиная с 1)")]     /**/ int row,
            [ExcelArgument(Description = "Столбец (начиная с 1)")]    /**/ int col,
            [ExcelArgument(Description = "Значение")]                 /**/ object varValue,
            [ExcelArgument(Description = "Необязательное возвращаемое значение вместо сохраняемого массива")] object returnValue = null)
        {
            //извлечь указанный массив
            dynamic array = objArr[varNumber];

            array[row - 1, col - 1] = varValue;

            if (returnValue == null || returnValue is ExcelMissing)
                return varValue;
            else
                return returnValue;
        }


        [ExcelFunction(Description = "Получить значение элемента двумерного массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrItemGet(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] int varNumber,
            [ExcelArgument(Description = "Строка (начиная с 1)")] /**/int row,
            [ExcelArgument(Description = "Столбец (начиная с 1)")]/**/int col)
        {
            //извлечь нужный массив
            dynamic array = objArr[varNumber];

            return array[row - 1, col - 1];
        }


        [ExcelFunction(Description = "Получить строку массива", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrGetRow(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15) ")] int varNumber,
            [ExcelArgument(Description = "Номер строки, начиная с 1")] int rowNumber)
        {
            //извлечь нужный массив
            object oArr = objArr[varNumber];

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


        [ExcelFunction(Description = "Транспонировать массив и сохранить его в одну из 16 переменных", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object ArrSetTransp(
        //--------------------------------------------------------------------------------------------------------------
            [ExcelArgument(Description = "Номер массива (от 0..15)")] /**/ int varNumber,
            [ExcelArgument(Description = "Массив")]                   /**/ object array,
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

                    //если элименты - простые типы значений, заполнить значениями можно исп. Buffer.BlockCopy
                    if (elementType.IsValueType)
                    {
                        int elSize = Marshal.SizeOf(elementType);

                        destArr = Array.CreateInstance(elementType, 1, srcArr.Length);        //размерности зад-ся с нуля
                        Buffer.BlockCopy(srcArr, 0, destArr, 0, (srcArr.Length) * elSize);

                        //Array.Copy(srcArr, 0, destArr, 0, srcArr.Length - 1)		//не работает, т.к.разное кол-во размерностей
                    }
                    else
                    {// в массиве элементы ссылочного типа, заполнять придётся вручную

                        dynamic txtSrcArr = array;
                        object[,] txtDstArr = new object[1, srcArr.Length];

                        for (int i = 0; i <= srcArr.Length - 1; i++)
                        {
                            txtDstArr[0, i] = txtSrcArr[i];
                        }
                        destArr = txtDstArr;
                    }
                }
                else if (srcArr.Rank == 2)
                // транспонируемый массив - двумерный
                {

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
                        // массив - один столбец
                        {
                            destArr = Array.CreateInstance(elementType, 1, srcArr.Length);    //размерности зад-ся с нуля
                        }
                        Array.Copy(srcArr, 0, destArr, 0, srcArr.Length);
                    }
                    else
                    // массив - прямоугольная матрица, будет цикл для преобразования
                    {

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
                // это либо не массив, либо размерностей больше 2
                {
                    return XlCVError.xlErrValue;
                }

                objArr[varNumber] = destArr;
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

        #endregion


        [ExcelFunction(Description = "Выполняет все функции в первом аргументе, а значение возвращет из второго", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static object Macro(
        /*------------------------------------------------------------------------------------------------------------*/
            [ExcelArgument(Description = "Тело макро функции")] object FormuLa,
            [ExcelArgument(Description = "Формула для вывода результата")] object Result = null)
        {
            if (Result is ExcelMissing)
                return ((Application)ExcelDnaUtil.Application).Caller.FormulaLocal();
            else
            {
                return Result;
            }
        }

        [ExcelFunction(Description = "My first .NET function", Category = "ANik")]
        //--------------------------------------------------------------------------------------------------------------
        public static string HelloR([ExcelArgument(Description = "Имя кого поприветствовать")] string Name)
        //--------------------------------------------------------------------------------------------------------------
        {
            return "Hello " + Name;
        }
    }
}
