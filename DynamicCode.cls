VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (dst_, src_) As Long
Private Declare PtrSafe Function VariantCopyInd Lib "oleaut32" (dst_, src_) As Long
Public Captures As Variant

Public Function Func43026x9926273148x0(xs, x) As Variant
'xs,x=>>$'{typename(xs)},{xs.count}:{x+1}''
    VariantCopyInd Func43026x9926273148x0, "" & (TypeName(xs)) & "," & (xs.Count) & ":" & (x + 1) & ""
End Function

Public Function Func43026x9923726852x0(xs, x) As Variant
'xs,x=>>$'{typename(xs)},{xs.count}:{(x+1)}''
    VariantCopyInd Func43026x9923726852x0, "" & (TypeName(xs)) & "," & (xs.Count) & ":" & ((x + 1)) & ""
End Function

Public Function Func43026x9738194444x0(xs, x) As Variant
'xs,x=>>x+1'
    VariantCopyInd Func43026x9738194444x0, x + 1
End Function

Private Function dict43026x9294675926x2(ay, x, y, z) As Dictionary: Set dict43026x9294675926x2 = New Dictionary: dict43026x9294675926x2.Add "ay", ay: dict43026x9294675926x2.Add "x", x: dict43026x9294675926x2.Add "y", y: dict43026x9294675926x2.Add "z", z: End Function

Private Function coll43026x9294675926x1(ay, x) As Collection: Set coll43026x9294675926x1 = New Collection: coll43026x9294675926x1.Add ay, "ay": coll43026x9294675926x1.Add x, "x": End Function

Public Function Func43026x9294675926x0(x) As Variant
'x=>>{x=x mod 3, y=[x], z=2}'
    VariantCopyInd Func43026x9294675926x0, dict43026x9294675926x2(Grouping.AnonymousIdentify, x Mod 3, coll43026x9294675926x1(Grouping.AnonymousIdentify, x), 2)
End Function

Private Function dict43026x9286805556x1(ay, x, y, z) As Dictionary: Set dict43026x9286805556x1 = New Dictionary: dict43026x9286805556x1.Add "ay", ay: dict43026x9286805556x1.Add "x", x: dict43026x9286805556x1.Add "y", y: dict43026x9286805556x1.Add "z", z: End Function

Public Function Func43026x9286805556x0(x) As Variant
'x=>>{x=x mod 3, y=new collection, z=2}'
    VariantCopyInd Func43026x9286805556x0, dict43026x9286805556x1(Grouping.AnonymousIdentify, x Mod 3, New Collection, 2)
End Function

Public Function Func43026x8969328704x0(x) As Variant
'x=>>x'
    VariantCopyInd Func43026x8969328704x0, x
End Function

Public Function Func43025x9873032407x0(x) As Variant
'x=>>$'x={x}''
    VariantCopyInd Func43025x9873032407x0, "x=" & x & ""
End Function

Public Function Func43025x9701388889x0(x, y) As Variant
'x,y=>>x+y'
    VariantCopyInd Func43025x9701388889x0, x + y
End Function

Public Function Func43023x7329398148x0(x) As Variant
'x : c, cc =>> x + c + cc+1'
    Dim c: VariantCopy c, Captures(0)
    Dim cc: VariantCopy cc, Captures(1)
    VariantCopyInd Func43023x7329398148x0, x + c + cc + 1
End Function

Public Function Func43023x7326851852x3(x, y) As Variant
'x,y=>>$'{grouping.SeriarizeIfAnonymous(x)}:{y}''
    VariantCopyInd Func43023x7326851852x3, "" & Grouping.SeriarizeIfAnonymous(x) & ":" & y & ""
End Function

Public Function Func43023x7326736111x2(x) As Variant
'x=>>$'[{x}]''
    VariantCopyInd Func43023x7326736111x2, "[" & x & "]"
End Function

Private Function dict43023x732662037x1(ay, x, y, z) As Dictionary: Set dict43023x732662037x1 = New Dictionary: dict43023x732662037x1.Add "ay", ay: dict43023x732662037x1.Add "x", x: dict43023x732662037x1.Add "y", y: dict43023x732662037x1.Add "z", z: End Function

Public Function Func43023x732662037x0(x) As Variant
'x=>>{x=x mod 3, y=1, z=2}'
    VariantCopyInd Func43023x732662037x0, dict43023x732662037x1(Grouping.AnonymousIdentify, x Mod 3, 1, 2)
End Function

Public Function Func43023x7321875x3(x) As Variant
'x=>>x+1'
    VariantCopyInd Func43023x7321875x3, x + 1
End Function

Public Function Func43023x7321875x2(x) As Variant
'x=>>x!x'
    VariantCopyInd Func43023x7321875x2, x!x
End Function

Private Function coll43023x7321759259x1(ay, x, y) As Collection: Set coll43023x7321759259x1 = New Collection: coll43023x7321759259x1.Add ay, "ay": coll43023x7321759259x1.Add x, "x": coll43023x7321759259x1.Add y, "y": End Function

Public Function Func43023x7321759259x0(x) As Variant
'x=>>[x,y='abc']'
    VariantCopyInd Func43023x7321759259x0, coll43023x7321759259x1(Grouping.AnonymousIdentify, x, "abc")
End Function

Public Function Func43023x7319675926x0(x) As Variant
'x=>>$'「{x}」''
    VariantCopyInd Func43023x7319675926x0, "「" & x & "」"
End Function

Private Function dict43023x7283680556x1(ay, a, b) As Dictionary: Set dict43023x7283680556x1 = New Dictionary: dict43023x7283680556x1.Add "ay", ay: dict43023x7283680556x1.Add "a", a: dict43023x7283680556x1.Add "b", b: End Function

Public Function Func43023x7283680556x0(x) As Variant
'x=>>{a=1,b=1}'
    VariantCopyInd Func43023x7283680556x0, dict43023x7283680556x1(Grouping.AnonymousIdentify, 1, 1)
End Function

Public Function Func43023x7281365741x0(x) As Variant
'x:c=>>x & c'
    Dim c: VariantCopy c, Captures(0)
    VariantCopyInd Func43023x7281365741x0, x & c
End Function

Public Function Instantiate() As Object: Set Instantiate = New DynamicCode: End Function

