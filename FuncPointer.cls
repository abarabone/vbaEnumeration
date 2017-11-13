VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FuncPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Binary

Implements IFunc



'＜現状＞

'　引数・戻り値はすべて variant でなくてはならない。

'　引数は ByRef/ByVal に対応。?

'　function のみ対応。sub は未対応。

'　いずれは戻り値/引数の数と型を保持して適応外で落ちないようにしたいが無理臭いかな…。




#If Win64 Then
    Const cSizeOfVariant& = 2 + 2 * 3 + 8 * 2   '24
    Const cSizeOfPointer& = 8
    Const cNullPointer^ = 0
#Else
    Const cSizeOfVariant& = 2 + 2 * 3 + 4 * 2   '16
    Const cSizeOfPointer& = 4
    Const cNullPointer& = 0
#End If




' win API ---------

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance_ As LongPtr, ByVal oVft_ As LongPtr, ByVal cc_ As Long, _
    ByVal vtReturn_ As Integer, _
    ByVal cActuals_ As Long, valueTypeTop_ As Integer, argPtrTop_ As LongPtr, _
    pvargResult_ As Variant _
) As Long

Const cStdCall& = 4


Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst_ As Any, ByRef src_ As Any, ByVal size_&)

'--------------------





'構造体定義 ============================================================



' Variant 内部レイアウト

Private Type VariantStruct
    varType     As Integer
    reserve0    As Integer
    reserve1    As Integer
    reserve2    As Integer
    pEntity0    As LongPtr
    pEntity1    As LongPtr
End Type





'メンバ宣言 ============================================================



Private pFunc_ As LongPtr '関数へのアドレス。

'アドレスは 0 でも安全（ DispCallFunc() はそのまま抜ける）。

    




'関数定義 ============================================================



' インスタンス生成／初期化 ---------------------------------


'オブジェクトインスタンスを新規生成する。

Public Function CNew(funcPointer_ As LongPtr) As FuncPointer
    
    Set CNew = New FuncPointer
    
    CNew.SetFunc funcPointer_
    
End Function


'関数アドレスをセットする。

Public Function SetFunc(funcPointer_ As LongPtr) As FuncPointer
    Set SetFunc = Me
    
    pFunc_ = funcPointer_
    
End Function


' --------------------------------------------------------------





'キャスト --------------------------------------------

Public Function AsIFunc() As IFunc
    
    Set AsIFunc = Me
    
End Function

' ----------------------------------------------------







'関数呼び出し -----------------------------------------------------------------------



'現状は配列間の隙間が発生するため使えない　未完成
Private Function IFunc_Exec(ParamArray args_()) As Variant
    
    If IsMissing(args_) Then
        
        DispCallFunc 0, pFunc_, cStdCall, vbVariant, 0, 0, 0, IFunc_Exec
        
        Exit Function
        
    End If
    
    
    Dim pSafeArray_ As LongPtr:     pSafeArray_ = Not Not args_
'    Dim pArgTop_ As LongPtr:        SafeArrayAccessData ByVal pSafeArray_, pArgTop_
'   Dim pArgTop_ As LongPtr:        pArgTop_ = xCom.GetAddressOfSafeArray(Not Not args_)
    
    Static types_%(16 - 1)
    Static pArgs_(16 - 1) As LongPtr
'    If (Not types_) = -1 Then initParams_ 7, types_, pArgs_
'    If UBound(types_) > UBound(args_) Then initParams_ UBound(args_) + 1, types_, pArgs_
    '動的配列にすると重くなるっぽい
    
    Dim i&
    For i = 0 To UBound(args_)
        
        types_(i) = vbVariant Or &H4000
        
'        pArgs_(i) = pArgTop_ + i * cSizeOfVariant
        
    Next
    
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, UBound(args_) + 1, types_(0), pArgs_(0), IFunc_Exec
    
    
'    SafeArrayUnaccessData ByVal pSafeArray_
    
    
End Function



Private Function IFunc_ExecMax06(Optional arg0_, Optional arg1_, Optional arg2_, Optional arg3_, Optional arg4_, Optional arg5_) As Variant
    
    
    Dim i&
    
    Select Case True
        Case IsMissing(arg0_):  i = 0
        Case IsMissing(arg1_):  i = 1
        Case IsMissing(arg2_):  i = 2
        Case IsMissing(arg3_):  i = 3
        Case IsMissing(arg4_):  i = 4
        Case IsMissing(arg5_):  i = 5
        Case Else:              i = 6
    End Select
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams i
    If UBound(xCom.CallableParamTypes) < i - 1 Then xCom.FormatCallableParams i
    
    On i GoTo On1, On2, On3, On4, On5
On6:    xCom.CallableParamArgs(5).pEntity0 = VarPtr(arg5_)
On5:    xCom.CallableParamArgs(4).pEntity0 = VarPtr(arg4_)
On4:    xCom.CallableParamArgs(3).pEntity0 = VarPtr(arg3_)
On3:    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
On2:    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
On1:    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, i, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), IFunc_ExecMax06
    
    
End Function






Private Function IFunc_xExec00() As Variant
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, 0, 0, 0, IFunc_xExec00
    
End Function

Private Function IFunc_xExec01(arg0_) As Variant
    
    Dim type_%
    Dim pArg_ As LongPtr
    Dim arg_ As VariantStruct
    
    type_ = vbVariant Or &H4000
    pArg_ = VarPtr(arg_)
    arg_.varType = vbVariant Or &H4000
    arg_.pEntity0 = VarPtr(arg0_)
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, 1, type_, pArg_, IFunc_xExec01
    
End Function

Private Function IFunc_xExec02(arg0_, arg1_) As Variant
    
    Const cParamLength_& = 2
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), IFunc_xExec02
    
End Function

Private Function IFunc_xExec03(arg0_, arg1_, arg2_) As Variant
    
    Const cParamLength_& = 3
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), IFunc_xExec03
    
End Function

Private Function IFunc_xExec04(arg0_, arg1_, arg2_, arg3_) As Variant
    
    Const cParamLength_& = 4
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
    xCom.CallableParamArgs(3).pEntity0 = VarPtr(arg3_)
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), IFunc_xExec04
    
End Function

Private Function IFunc_xExec05(arg0_, arg1_, arg2_, arg3_, arg4_) As Variant
    
    Const cParamLength_& = 5
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
    xCom.CallableParamArgs(3).pEntity0 = VarPtr(arg3_)
    xCom.CallableParamArgs(4).pEntity0 = VarPtr(arg4_)
    
    DispCallFunc 0, pFunc_, cStdCall, vbVariant, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), IFunc_xExec05
    
End Function


' ----------------------------------------------




