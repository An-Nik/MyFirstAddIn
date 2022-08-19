using static MacroFunctions.SetGetVars;
using System.Runtime.InteropServices;
using System;

object q;

Byte[,] Arr0 = { { 1, 2, 3 }, { 4, 5, 6 } };
SetArray(0, Arr0);


Int64[] Arr1 = { 1, 2, 3 };
Arr2SetTransp(0, Arr1);
q = GetArray(0);


Double[,] Arr2 = { { 1 }, { 2 }, { 3 } };
Arr2SetTransp(0, Arr2);
q = GetArray(0);


Byte[,] Arr3 = { { 1, 2, 3 }, { 4, 5, 6 } };
Arr2SetTransp(0, Arr3);
q = GetArray(0);


string[,] Arr4 = { { "one" }, { "tw" }, { "tree" }, { "four" } };
Arr2SetTransp(0, Arr4);
q = GetArray(0);


string[] Arr5 = { "one", "two", "tree", "four" };
Arr2SetTransp(0, Arr5);
q = GetArray(0);


string[,] Arr = { { "one", "tw" }, { "tree", "four" } };
Arr2SetTransp(0, Arr);
q = GetArray(0);

q = GetArray(0, 1, 2);



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

    SetArray(0, Arr1);
    var q1 = GetArray(0, 1, 2);
    
}
object o = (Int32)24;
int? i = o as int?;

long d = 2;
int iD = Convert.ToInt32(d);

Console.WriteLine("The double value {0} when converted to an int becomes {1}", d, iD);
