VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Delegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFunc



' win API ---------

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance_ As LongPtr, ByVal oVft_ As LongPtr, ByVal cc_ As Long, _
    ByVal vtReturn_ As Integer, _
    ByVal cActuals_ As Long, valueTypeTeop_ As Integer, argPtrTeop_ As LongPtr, _
    pvargResult_ As Variant _
) As Long

Private Const cStdCall& = 4


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




Private obj_            As Object   ' LontPtr で持つと、開放される恐れがある。

Private methodOffset_   As Long     ' V-table offset (byte size)







'関数定義 ============================================================


' インスタンス生成 ---------------------------------


'オブジェクトインスタンスを新規生成する。

Public Function CNew(Optional object_ As Object, Optional methodName_) As Delegate
    Set CNew = New Delegate
    
    CNew.SetMethod object_, methodName_
    
End Function



'複製を作って返す。メンバは浅いコピーとなる。

Public Function Clone() As Delegate
    
    Set Clone = Me.CNew.SetMethod(obj_, methodOffset_)
    
End Function


' --------------------------------------------------





'初期化 --------------------------------------------


'オブジェクトとメソッドを指定する。
'　指定した引数だけがメンバを更新される。

Public Function SetMethod(Optional object_ As Object, Optional methodName_) As Delegate
    Set SetMethod = Me
    
    If Not object_ Is Nothing Then _
        Set obj_ = object_
    
    If Not IsMissing(methodName_) Then
        
        If IsNumeric(methodName_) Then _
            methodOffset_ = methodName_
        
        If varType(methodName_) = vbString Then _
            methodOffset_ = xCom.GetOrderByName(obj_, CStr(methodName_))
        
        'If methodName_ = -1 Then xxx'エラー処理
        
    End If
    
End Function


' --------------------------------------------------








'情報取得 -----------------------------------


'対象オブジェクトを返す。

Public Property Get Target() As Object
    
    Set Target = obj_
    
End Property



'メソッド名確認

Public Function HasMethod(methodName_$) As Boolean
    
    If methodOffset_ = -1 Then Exit Function
    
    
    Dim ofs_&:  ofs_ = xCom.GetOrderByName(obj_, methodName_)
    
    If ofs_ = -1 Then Exit Function
    
    
    HasMethod = (methodOffset_ = ofs_)
    
End Function



' --------------------------------------------------------------





'キャスト --------------------------------------------

Public Function AsIFunc() As IFunc
    
    Set AsIFunc = Me
    
End Function

' ----------------------------------------------------






'関数呼び出し -----------------------------------------------------------------------



'現状は配列間の隙間が発生するため使えない 未完成

Private Function IFunc_Exec(ParamArray args_()) As Variant
    
    If IsMissing(args_) Then
        
        MoveMemory IFunc_Exec, IFunc_xExec00(), LenB(IFunc_Exec)
        
        Exit Function
        
    End If
    
    
    Dim v_ As VariantStruct
    v_.varType = vbVariant Or &H4000 Or &H2000
    v_.pEntity0 = Not Not args_
    
    Dim vArgs_
    MoveMemory vArgs_, v_, LenB(v_)
    
    xCom.CallFunc2 ObjPtr(obj_), methodOffset_, vArgs_, out_result_:=IFunc_Exec, isVbaClass_:=True
    
    MoveMemory vArgs_, Empty, LenB(v_)
    
    Exit Function
    
    Dim pSafeArray_ As LongPtr
    pSafeArray_ = Not Not args_
    
    Dim pArgs_ As LongPtr
'    SafeArrayAccessData ByVal pSafeArray_, pArgs_
'    pArgs_ = xCom.GetAddressOfSafeArray(Not Not args_)
    
    xCom.CallFunc ObjPtr(obj_), methodOffset_, pArgs_, UBound(args_) + 1, out_result_:=IFunc_Exec, isVbaClass_:=True
    
'    SafeArrayUnaccessData ByVal pSafeArray_
    
    
End Function


Private Function IFunc_ExecMax06( _
 _
    Optional arg0_, Optional arg1_, Optional arg2_, Optional arg3_, Optional arg4_, Optional arg5_ _
 _
) As Variant
    
    
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
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams i + 1
    If UBound(xCom.CallableParamTypes) < i Then xCom.FormatCallableParams i + 1
    
    On i GoTo On1, On2, On3, On4, On5
On6:    xCom.CallableParamArgs(5).pEntity0 = VarPtr(arg5_)
On5:    xCom.CallableParamArgs(4).pEntity0 = VarPtr(arg4_)
On4:    xCom.CallableParamArgs(3).pEntity0 = VarPtr(arg3_)
On3:    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
On2:    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
On1:    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    
    xCom.CallableParamArgs(i).pEntity0 = VarPtr(IFunc_ExecMax06)
    
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, i + 1, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), resForCom_
    
    
End Function






Private Function IFunc_xExec00() As Variant
    
    Dim type_%
    Dim pArg_ As LongPtr
    Dim arg_ As VariantStruct
    
    arg_.varType = vbVariant Or &H4000
    arg_.pEntity0 = VarPtr(IFunc_xExec00)
    
    type_ = arg_.varType
    pArg_ = VarPtr(arg_)
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, 1, type_, pArg_, resForCom_
    
End Function

Private Function IFunc_xExec01(arg0_) As Variant
    
    Const cParamLength_& = 1 + 1
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(IFunc_xExec01)
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), resForCom_
    
End Function

Private Function IFunc_xExec02(arg0_, arg1_) As Variant
    
    Const cParamLength_& = 2 + 1
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(IFunc_xExec02)
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), resForCom_
    
End Function

Private Function IFunc_xExec03(arg0_, arg1_, arg2_) As Variant
    
    Const cParamLength_& = 3 + 1
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
    xCom.CallableParamArgs(3).pEntity0 = VarPtr(IFunc_xExec03)
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), resForCom_
    
End Function

Private Function IFunc_xExec04(arg0_, arg1_, arg2_, arg3_) As Variant
    
    Const cParamLength_& = 4 + 1
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
    xCom.CallableParamArgs(3).pEntity0 = VarPtr(arg3_)
    xCom.CallableParamArgs(4).pEntity0 = VarPtr(IFunc_xExec04)
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), resForCom_
    
End Function

Private Function IFunc_xExec05(arg0_, arg1_, arg2_, arg3_, arg4_) As Variant
    
    Const cParamLength_& = 5 + 1
    
    If (Not xCom.CallableParamTypes) = -1 Then xCom.FormatCallableParams cParamLength_
    If UBound(xCom.CallableParamTypes) < cParamLength_ - 1 Then xCom.FormatCallableParams cParamLength_
    
    xCom.CallableParamArgs(0).pEntity0 = VarPtr(arg0_)
    xCom.CallableParamArgs(1).pEntity0 = VarPtr(arg1_)
    xCom.CallableParamArgs(2).pEntity0 = VarPtr(arg2_)
    xCom.CallableParamArgs(3).pEntity0 = VarPtr(arg3_)
    xCom.CallableParamArgs(4).pEntity0 = VarPtr(arg4_)
    xCom.CallableParamArgs(5).pEntity0 = VarPtr(IFunc_xExec05)
    
    Dim resForCom_
    
    DispCallFunc ObjPtr(obj_), methodOffset_, cStdCall, vbLong, cParamLength_, xCom.CallableParamTypes(0), xCom.CallableParamPtrs(0), resForCom_
    
End Function


' ----------------------------------------------------------------------------

