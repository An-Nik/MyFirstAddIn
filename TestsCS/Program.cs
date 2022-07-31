using static MacroFunctions.SetGetVars;
using System.Runtime.InteropServices;
using System;

object q;

Int64[] Arr1 = { 1, 2, 3 };
SetTranspArr("q", Arr1);
q = GetArr("q");


Double[,] Arr2 = { { 1 }, { 2 }, { 3 } };
SetTranspArr("q", Arr2);
q = GetArr("q");


Byte[,] Arr3 = { { 1, 2, 3 }, { 4, 5, 6 } };
SetTranspArr("q", Arr3);
q = GetArr("q");


string[,] Arr4 = { { "one" }, { "tw" }, { "tree" }, { "four" } };
SetTranspArr("q", Arr4);
q = GetArr("q");


string[] Arr5 = { "one", "two", "tree", "four" };
SetTranspArr("q", Arr5);
q = GetArr("q");


string[,] Arr = { { "one", "tw" }, { "tree", "four" } };
SetTranspArr("q", Arr);
q = GetArr("q");

q = GetArr("q", 1, 2);



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

    SetArr("q", Arr1);
    var q1 = GetArr("q", (long)1, 2);
    
}
object o = (Int32)24;
int? i = o as int?;

long d = 2;
int iD = Convert.ToInt32(d);

Console.WriteLine("The double value {0} when converted to an int becomes {1}", d, iD);
