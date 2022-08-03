using static MacroFunctions.SetGetVars;
using System.Runtime.InteropServices;
using System;

object q;

ArrSet(1,new object[3]);
ArrItemSet(1, 0, 1);
ArrItemSet(1, 1, 2);
ArrItemSet(1, 2, 3);
q = ArrGet(1);


Int16[] Arr1 = { 1, 2, 3 };
Arr2SetTransp(1, Arr1);
q = ArrGet(1);


Double[,] Arr2 = { { 1.1 }, { 0.2 }, { -3.0 } };
Arr2SetTransp(1, Arr2);
q = ArrGet(1);


byte[,] Arr3 = { { 1, 2, 3 } };
Arr2SetTransp(1, Arr3);
q = ArrGet(1);


string[,] Arr4 = { { "one" }, { "tw" }, { "tree" }, { "four" } };
Arr2SetTransp(1, Arr4);
q = ArrGet(1);


string[] Arr5 = { "one", "two", "tree", "four" };
Arr2SetTransp(1, Arr5);
q = ArrGet(1);


object[,] Arr = { { "one", "tw" }, { "tree", "four" } };
Arr2SetTransp(1, Arr);
q = ArrGet(1);

q = ArrGet(1, 1, 2);


{
    String[,] strArr = { { "1" }, { "2" } };

    object oArr = strArr;

    dynamic dArr = strArr;

    var aArr = (Array)dArr;

    string[,] tmpStrArr = dArr;

    object o2 = dArr[0, 0];

    string s = tmpStrArr[1, 0];
}

{
    object a = Arr1;
    object b = Marshal.SizeOf(Arr1.GetType().GetElementType());

    ArrSet(1, Arr1);
    var q1 = ArrGet(1, (long)1, 2);
    
}
object o = (Int32)24;
int? i = o as int?;

long d = 2;
int iD = Convert.ToInt32(d);

Console.WriteLine("The double value {0} when converted to an int becomes {1}", d, iD);
