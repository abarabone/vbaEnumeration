Attribute VB_Name = "xCom"
Option Explicit





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


Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (ByRef dst_ As Any, ByRef src_ As Any) As Long

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst_ As Any, ByRef src_ As Any, ByVal size_&)


Private Declare PtrSafe Function CoTaskMemAlloc Lib "Ole32" (ByVal byte_&) As LongPtr

Private Declare PtrSafe Sub CoTaskMemFree Lib "Ole32" (ByVal pMem_ As LongPtr)

'--------------------





Private Const cE_NOINTERFACE& = &H80004002






'構造体定義 ============================================================



' GUID 格納用構造体（中身はいじらないので１６バイトあればいい）

Private Type CoGuid
    id0         As Long
    id1         As Integer
    id2         As Integer
    id3(8 - 1)  As Byte
End Type

'    // IID_NULL is null Interface ID, used when no other Interface ID is known.
'
'    IID_NULL = NewGUID("{00000000-0000-0000-0000-000000000000}")
'
'    // IID_IUnknown is for IUnknown interfaces.
'
'    IID_IUnknown = NewGUID("{00000000-0000-0000-C000-000000000046}")
'
'    // IID_IDispatch is for IDispatch interfaces.
'
'    IID_IDispatch = NewGUID("{00020400-0000-0000-C000-000000000046}")
'
'    // IID_IEnumVariant is for IEnumVariant interfaces
'
'    IID_IEnumVariant = NewGUID("{00020404-0000-0000-C000-000000000046}")





' TypeInfo オブジェクトの関数情報構造体

Private Type CoFuncDesc
    memid               As Long
    lprgscode           As LongPtr
    lprgelemdescParam   As LongPtr
    funckind            As Long
    invkind             As Long
    callconv            As Long
    cParams             As Integer
    cParamsOpt          As Integer
    oVft                As Integer
    cScodes             As Integer
    elemdescFunc        As LongPtr
    wFuncFlags          As Integer
End Type




' TypeInfo オブジェクトの属性情報構造体（サイズ短縮版）

Private Type CoTypeAttrLite
    Guid                As CoGuid
    lcid                As Long
    dwReserved          As Long
    memidConstructor    As Long
    memidDestructor     As Long
    lpstrSchema         As LongPtr
    cbSizeInstance      As Long
    typekind            As Long
    cFuncs              As Integer
    cVars               As Integer
    cImplTypes          As Integer
    cbSizeVft           As Integer
    cbAlignment         As Integer
    wTypeFlags          As Integer
    wMajorVerNum        As Integer
    wMinorVerNum        As Integer
'    tdescAlias          As TYPEDESC '14/26   72/88   ; (3*Ptr)+ 2 ? https://msdn.microsoft.com/en-us/library/aa911717.aspx
'    idldescType         As IDLDESC  '6       78/114  ; ? https://msdn.microsoft.com/en-us/library/aa909796.aspx
End Type





' Invoke 引数情報

Private Type DispParamsStruct
    pArgs_          As LongPtr
    pArgNames_      As LongPtr
    argLength_      As Long
    argNameLength_  As Long
End Type




' Variant 内部レイアウト

Private Type VariantStruct
    varType     As Integer
    reserve0    As Integer
    reserve1    As Integer
    reserve2    As Integer
    pEntity0    As LongPtr
    pEntity1    As LongPtr
End Type




'自前 EnumVariant 構造体

Public Type EnumVariantStruct
    
    PVtable         As LongPtr
    
    OperatorFunc    As IFunc
    
    RefCount        As Long
    
End Type




' SafeArray 内部レイアウト

Private Type SafeArrayStruct
    
    cDims       As Integer
    
    fFeatures   As Integer
    
    
    cbElements  As Long     '１要素のバイトサイズ
    
    cLocks      As Long     'ロック数。こちらでカウントアップしてしまうと、ＶＢＡ側から操作できなくなる。
    
    pvData      As LongPtr  'PVOID
    
    
    'SafeArrayBound rgsabound[ cDims ];
    
End Type
Private Type SafeArrayBound
    
    Elements    As Long     '要素数
    
    BaseIndex   As Long     ' LBound()
    
End Type






' 共用変数 ==========================================================


' CallDispFunc() 用の配列。
'　モジュールレベルで宣言しないで済む方法を模索したい。
'　　・ Static と同等の効果（一度行えば以降すむ）
'　　・ Redim できる
'　　・関数を通さないで済む（できれば）
'　　という条件を満たす方法はなさそうかなぁ…。CoTaskMem() とか？

Public CallableParamTypes() As Integer
Public CallableParamPtrs()  As LongPtr
Public CallableParamArgs()  As VariantStruct








'ユーティリティ関数 =======================================================



'ＧＵＩＤを文字列に変換する。

Public Function GuidToString(guid_ As CoGuid) As String
    
    Dim ts_$()
    ReDim ts_(5 - 1)
    
    ts_(0) = Right("00000000" & Hex(guid_.id0), 8)
    ts_(1) = Right("0000" & Hex(guid_.id1), 4)
    ts_(2) = Right("0000" & Hex(guid_.id2), 4)
    ts_(3) = "0000"
    ts_(4) = "000000000000"
    
    Dim i&, istr_&
    
    For i = 1 To 4 - 1 Step 2
        Mid(ts_(3), i, 2) = Right("00" & Hex(guid_.id3(istr_)), 2)
        istr_ = istr_ + 1
    Next
    
    For i = 1 To 12 - 1 Step 2
        Mid(ts_(4), i, 2) = Right("00" & Hex(guid_.id3(istr_)), 2)
        istr_ = istr_ + 1
    Next
    
    GuidToString = Join(ts_, "-")
    
End Function

Private Function isIUnknown_(guid_ As CoGuid) As Boolean
    
    If Not (guid_.id0 = 0 And guid_.id1 = 0 And guid_.id2 = 0) Then Exit Function
    
    Dim bs_(8 - 1) As Byte: bs_(0) = &HC0: bs_(7) = &H46
    
    isIUnknown_ = CStr(guid_.id3) = CStr(bs_)
    
End Function

Private Function isIEnumVariant_(guid_ As CoGuid) As Boolean
    
    If Not (guid_.id0 = &H20404 And guid_.id1 = 0 And guid_.id2 = 0) Then Exit Function
    
    Dim bs_(8 - 1) As Byte: bs_(0) = &HC0: bs_(7) = &H46
    
    isIEnumVariant_ = CStr(guid_.id3) = CStr(bs_)
    
End Function




' variant 値を受け取り、それへの参照 variant 値を生成して返す。
'参照元が破棄された場合、アクセスすると落ちるので注意。

Public Function ref(value_) As Variant
    
    Dim v_ As VariantStruct
    
    MoveMemory v_, value_, LenB(v_)
    
    
    If (v_.varType And &H4000) = 0 Then
        
        v_.varType = v_.varType Or &H4000
        
        v_.pEntity0 = VarPtr(value_) + (VarPtr(v_.pEntity0) - VarPtr(v_))
        
    End If
    
    
    MoveMemory ref, v_, LenB(v_)
    
End Function





'受け取った variant 値を Empty にして、コピーした variant を戻り値として返す。
'　※戻り値を別の variant に代入した場合は、配列のコピーは行われない（らしい、検証したがいまいちわからんかった）。

Public Function Move(value_) As Variant
    
    MoveMemory Move, value_, cSizeOfVariant
    
    MoveMemory value_, Empty, cSizeOfVariant '元を空にしないと開放が２重に行われて落ちる。
    
End Function







'調整中

Public Sub CallFunc2(pObj_ As LongPtr, pFunc_ As LongPtr, ByRef args_, Optional ByRef out_result_, Optional isVbaClass_ As Boolean)
    
    Dim argLength_&:    argLength_ = UBound(args_) + 1
    
    Dim argResLength_&: argResLength_ = argLength_ + (Not IsMissing(out_result_)) * isVbaClass_
    
    
    
    '引数用配列の初期化
    
    Static types_%()
    Static pArgs_() As LongPtr
    
    Select Case True
        
        Case (Not types_) = -1 '未初期化
            
            Dim minLen_&: minLen_ = IIf(argResLength_ > 7, argResLength_, 7) '最低 10 は確保される。
            
            initArgArray_ minLen_, types_, pArgs_
            
        Case argResLength_ > UBound(types_) + 1
            
            initArgArray_ argResLength_, types_, pArgs_
            
    End Select
    
    
    '引数のセット
    
    Dim i&
    For i = 0 To argLength_ - 1
        
        pArgs_(i) = VarPtr(args_(i))
        
    Next
    
    
    
    '関数をコール
    
    If Not isVbaClass_ Then
        
        Dim resultType_ As VbVarType
        
        If Not IsMissing(out_result_) Then resultType_ = vbVariant
        
        DispCallFunc pObj_, pFunc_, cStdCall, resultType_, argLength_, types_(0), pArgs_(0), out_result_
        
    Else
        
        If argResLength_ > argLength_ Then
            
            Dim result_
            
            Dim resultArg_ As VariantStruct
            resultArg_.varType = vbVariant Or &H4000
            resultArg_.pEntity0 = VarPtr(result_)
            
            types_(argLength_) = vbVariant Or &H4000
            pArgs_(argLength_) = VarPtr(resultArg_)
            
        End If
        
        Dim res_
        
        DispCallFunc pObj_, pFunc_, cStdCall, vbLong, argResLength_, types_(0), pArgs_(0), res_
        
    End If
    
    
End Sub

Public Sub CallFunc( _
 _
    pObj_ As LongPtr, pFunc_ As LongPtr, pArgTop_ As LongPtr, argLength_&, _
 _
    Optional ByRef out_result_, Optional isVbaClass_ As Boolean _
 _
)
    
    Dim argResLength_&: argResLength_ = argLength_ + (Not IsMissing(out_result_)) * isVbaClass_
    
    
    
    '引数用配列の初期化
    
    Static types_%()
    Static pArgs_() As LongPtr
    
    Select Case True
        
        Case (Not types_) = -1 '未初期化
            
            Dim minLen_&: minLen_ = IIf(argResLength_ > 7, argResLength_, 7) '最低 10 は確保される。
            
            initArgArray_ minLen_, types_, pArgs_
            
        Case argResLength_ > UBound(types_) + 1
            
            initArgArray_ argResLength_, types_, pArgs_
            
    End Select
    
    
    '引数のセット
    
    Dim i&
    For i = 0 To argLength_ - 1
        
        pArgs_(i) = pArgTop_ + i * cSizeOfVariant
        
    Next
    
    
    
    '関数をコール
    
    If Not isVbaClass_ Then
        
        Dim resultType_ As VbVarType
        
        If Not IsMissing(out_result_) Then resultType_ = vbVariant
        
        DispCallFunc pObj_, pFunc_, cStdCall, resultType_, argLength_, types_(0), pArgs_(0), out_result_
        
    Else
        
        If argResLength_ > argLength_ Then
            
            Dim result_
            
            Dim resultArg_ As VariantStruct
            resultArg_.varType = vbVariant Or &H4000
            resultArg_.pEntity0 = VarPtr(result_)
            
            types_(argLength_) = vbVariant Or &H4000
            pArgs_(argLength_) = VarPtr(resultArg_)
            
        End If
        
        Dim res_
        
        DispCallFunc pObj_, pFunc_, cStdCall, vbLong, argResLength_, types_(0), pArgs_(0), res_
        
    End If
    
    
End Sub

'引数用配列を初期化する。

Private Sub initArgArray_(length_&, ByRef types_%(), ByRef pArgs_() As LongPtr)
    
    Dim capacity_&: capacity_ = length_ + length_ \ 2 - 1 '必要数の 1.5 倍を確保。
    
    ReDim types_(capacity_)
    ReDim pArgs_(capacity_)
    
    Dim i&
    For i = 0 To capacity_
        
        types_(i) = vbVariant Or &H4000
        
    Next
    
End Sub

'－－－－－－－－－－－－－－－－－－









' dispId からメソッドを呼び出す。

Public Function InvokeMethod(object_ As Object, dispId_&, ParamArray args_()) As Variant
    
    Call_Invoke object_, dispId_, VarPtr(args_(0)), UBound(args_) + 1, out_result_:=InvokeMethod
    
End Function










'メソッド名前から序数を取得する

Public Function GetOrderByName(object_ As Object, methodName_$) As Long
    
    Dim typeInfo_ As IUnknown
    Set typeInfo_ = Call_GetTypeInfo(object_)
    
    Dim dispId_&:   dispId_ = Call_GetIDsOfNames(object_, methodName_)
    
    
    Dim i&
    For i = 0 To call_GetTypeAttr_(typeInfo_).cFuncs - 1
        
        ' funcDesc 構造体を取得して dispId が一致するメソッドを探し、vtable 上の順番を取得する。
        
        Dim desc_ As CoFuncDesc
        desc_ = call_GetFuncDesc_(typeInfo_, i)
        
        If desc_.memid = dispId_ Then
            
            GetOrderByName = desc_.oVft
            
            Exit Function
            
        End If
        
    Next
    
    GetOrderByName = -1
    
End Function







'オブジェクトメンバメソッド名リストの取得

Public Function MakeVtableList(object_ As Object) As Collection
    
    Dim typeInfo_ As IUnknown
    Set typeInfo_ = Call_GetTypeInfo(object_)
    
    
    Dim c_ As New Collection
    
    
    Dim i&
    For i = 0 To call_GetTypeAttr_(typeInfo_).cFuncs - 1
        
        
        ' funcDesc 構造体を取得して、memid と vtable 上の順番を取得する。
        
        Dim desc_ As CoFuncDesc
        desc_ = call_GetFuncDesc_(typeInfo_, i)
        
        
        ' dispId_ からメソッドの名前を取得する。
        
        Dim thisName_$
        thisName_ = call_GetNames(typeInfo_, desc_.memid)
        
        
        c_.Add desc_.oVft, thisName_
        
    Next
    
    
    Set MakeVtableList = c_
    
End Function








'メソッドアドレスを取得する。

Public Function GetAddress(object_ As Object, methodName_$) As LongPtr
    
    Dim typeInfo_ As IUnknown
    Set typeInfo_ = Call_GetTypeInfo(object_)
    
    
    Dim dispId_&
    dispId_ = Call_GetIDsOfNames(object_, methodName_)
    
    
    GetAddress = Call_AddressOfMember(typeInfo_, dispId_)
    
End Function



' v-table からメソッドアドレスを取得する。

Public Function GetAddressFromVtable(object_ As Object, methodName_$) As LongPtr
'    Dim ti_: Set ti_ = call_GetTypeInfo_(object_)
    Dim order_&:    order_ = xCom.GetOrderByName(object_, methodName_)
    
    
    Dim pVTable_ As LongPtr
    
    MoveMemory ByVal VarPtr(pVTable_), ByVal ObjPtr(object_), LenB(pVTable_)
    
    MoveMemory ByVal VarPtr(GetAddressFromVtable), ByVal pVTable_ + order_, LenB(GetAddressFromVtable)
    
'    Dim a_(64) As LongPtr
'    MoveMemory ByVal VarPtr(a_(0)), ByVal pVTable_, 32 * 4
    
End Function











'インターフェースメソッドコール ===============================================================================




' TypeInfo 取得 ----------------------------

Public Function Call_GetTypeInfo(object_ As Object) As IUnknown
    
    
    'パラメータ設定
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForGetTypeInfo_ types_, ptrs_, _
        get_typeInfo_:=Call_GetTypeInfo
    
    
    
    '取得
    
    Dim res_
    
    Const cGetTypeInfo As LongPtr = 4 * cSizeOfPointer
    
    DispCallFunc ObjPtr(object_), cGetTypeInfo, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), res_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Function

' - - - - - - - -

Private Sub setParamsForGetTypeInfo_( _
 _
    ByRef ref_types_%(), ByRef ref_ptrs_() As LongPtr, _
 _
    ByRef get_typeInfo_ As IUnknown _
 _
)
    
    If (Not Not ref_types_) = 0 Then
        
        ReDim ref_types_(3 - 1)
        ReDim ref_ptrs_(3 - 1)
        
        '引数
        Static zero_:   zero_ = 0&                      ' 0 では Integer になるのでダメ
        Static locale_: locale_ = &H800&                ' LOCALE_SYSTEM_DEFAULT
        Static ppTi_:   ppTi_ = VarPtr(get_typeInfo_)   ' IUnknown 型のもつポインタ値を書き換えて取得する。
        
        ref_types_(0) = vbLong
        ref_ptrs_(0) = VarPtr(zero_)
        
        ref_types_(1) = vbLong
        ref_ptrs_(1) = VarPtr(locale_)
        
        ref_types_(2) = varType(ppTi_)
        ref_ptrs_(2) = VarPtr(ppTi_)
        
    End If
    
    
    ppTi_ = VarPtr(get_typeInfo_)
    
    
End Sub

' ----------------------------------------------------------------





' dispid 取得 ----------------------------

Public Function Call_GetIDsOfNames(object_ As Object, fname_$) As Long
    
    
    'パラメータ設定
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForGetDispId_ types_, ptrs_, _
        give_fname_:=fname_, _
        get_dispId_:=Call_GetIDsOfNames
    
    
    
    '取得
    
    Dim res_
    
    Const cGetIDsOfNames As LongPtr = 5 * cSizeOfPointer
    
    DispCallFunc ObjPtr(object_), cGetIDsOfNames, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), res_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Function

' - - - - - - - -

Private Sub setParamsForGetDispId_( _
 _
    ByRef ref_types_%(), ByRef ref_ptrs_() As LongPtr, _
 _
    give_fname_$, _
 _
    ByRef get_dispId_& _
 _
)
    
    If (Not Not ref_types_) = 0 Then
        
        ReDim ref_types_(5 - 1)
        ReDim ref_ptrs_(5 - 1)
        
        Static iid_ As CoGuid       'ＧＵＩＤはゼロ
        Static pName_ As LongPtr    '関数名
        
        '引数
        Static pIid_:   pIid_ = VarPtr(iid_)
        Static ppName_: ppName_ = VarPtr(pName_)
        Static count_:  count_ = 1&                 ' 1 では Integer となりバグる、このへんシビア
        Static locale_: locale_ = &H800&            ' LOCALE_SYSTEM_DEFAULT
        Static pId_:    pId_ = VarPtr(get_dispId_)
        
        ref_types_(0) = varType(pIid_)
        ref_ptrs_(0) = VarPtr(pIid_)
        
        ref_types_(1) = varType(ppName_)
        ref_ptrs_(1) = VarPtr(ppName_)
        
        ref_types_(2) = vbLong
        ref_ptrs_(2) = VarPtr(count_)
        
        ref_types_(3) = vbLong
        ref_ptrs_(3) = VarPtr(locale_)
        
        ref_types_(4) = varType(pId_)
        ref_ptrs_(4) = VarPtr(pId_)
        
    End If
    
    
    pName_ = StrPtr(give_fname_)
    
    pId_ = VarPtr(get_dispId_)
    
    
End Sub

' ----------------------------------------------------------------






'メソッド情報を取得、FuncDesc 構造体を返す ----------------------------

Private Function call_GetFuncDesc_(typeInfo_ As IUnknown, funcIndex_&) As CoFuncDesc
    
    
    'パラメータ設定
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    Dim pDesc_ As LongPtr '結果 FuncDesc の先頭
    
    setParamsForGetFuncDesc_ types_, ptrs_, _
        give_funcIndex_:=funcIndex_, _
        get_pDesc_:=pDesc_
    
    
    
    
    ' FuncDesc 構造体の取得
    
    Dim resA_
    
    Const cGetFuncDesc As LongPtr = 5 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cGetFuncDesc, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), resA_
    
    'if res_ =  then err.Raise 0,,""
    
    
    
    '構造体をコピーする
    
    MoveMemory ByVal VarPtr(call_GetFuncDesc_), ByVal pDesc_, LenB(call_GetFuncDesc_)
    
    
    
    ' FuncDesc 構造体の開放
    
    Dim resB_
    Dim pDescRelease_: pDescRelease_ = pDesc_
    
    Const cReleaseFuncDesc As LongPtr = 20 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cReleaseFuncDesc, cStdCall, vbEmpty, 1, varType(pDescRelease_), VarPtr(pDescRelease_), resB_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Function

' - - - - - - - -

Private Sub setParamsForGetFuncDesc_( _
 _
    ByRef ref_types_%(), ByRef ref_ptrs_() As LongPtr, _
 _
    give_funcIndex_&, _
 _
    get_pDesc_ As LongPtr _
 _
)
    
    If (Not Not ref_types_) = 0 Then
            
        ReDim ref_types_(2 - 1)
        ReDim ref_ptrs_(2 - 1)
        
        '引数
        Static fIndex_: fIndex_ = give_funcIndex_
        Static ppDesc_: ppDesc_ = VarPtr(get_pDesc_)
        
        ref_types_(0) = vbLong
        ref_ptrs_(0) = VarPtr(fIndex_)
        
        ref_types_(1) = varType(ppDesc_)
        ref_ptrs_(1) = VarPtr(ppDesc_)
        
    End If
    
    
    ppDesc_ = VarPtr(get_pDesc_)
    
    fIndex_ = give_funcIndex_
    
    
End Sub

' ----------------------------------------------------------------








'名前の取得 ----------------------------

Public Function call_GetNames(typeInfo_ As IUnknown, dispId_&) As String
    
    
    'パラメータ設定
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    Dim pName_ As LongPtr '受け取る名前文字列の先頭（なんで開放しなくていいんだろう？）
    
    setParamsForGetName_ types_, ptrs_, _
        give_dispId_:=dispId_, _
        get_pName_:=pName_
    
    
    
    '取得
    
    Dim res_
    
    Const cGetNames As LongPtr = 7 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cGetNames, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), res_ ' get names
    
    'if res_ =  then err.Raise 0,,""
    
    
    
    '結果を格納する。
    
    MoveMemory ByVal VarPtr(call_GetNames), ByVal VarPtr(pName_), LenB(pName_) ' pName_ は bstr
    
End Function

' - - - - - - - -

Private Sub setParamsForGetName_( _
 _
    ByRef ref_types_%(), ByRef ref_ptrs_() As LongPtr, _
 _
    give_dispId_&, _
 _
    get_pName_ As LongPtr _
 _
)
    
    If (Not Not ref_types_) = 0 Then
        
        ReDim ref_types_(4 - 1)
        ReDim ref_ptrs_(4 - 1)
        
        '参照先
        Static pRtnLength_& '取得数が返ってくるが、使用しない。
        
        '引数
        Static dispId_:     dispId_ = give_dispId_
        Static maxLength_:  maxLength_ = 1&                 ' 1 では Integer となりバグる、このへんシビア
        Static ppName_:     ppName_ = VarPtr(get_pName_)
        Static ppLength_:   ppLength_ = VarPtr(pRtnLength_)
        
        ref_types_(0) = vbLong
        ref_ptrs_(0) = VarPtr(dispId_)
        
        ref_types_(1) = varType(ppName_)
        ref_ptrs_(1) = VarPtr(ppName_)
        
        ref_types_(2) = vbLong
        ref_ptrs_(2) = VarPtr(maxLength_)
        
        ref_types_(3) = varType(ppLength_)
        ref_ptrs_(3) = VarPtr(ppLength_)
        
    End If
    
    
    dispId_ = give_dispId_
    
    ppName_ = VarPtr(get_pName_)
    
    
End Sub

' ----------------------------------------------------------------




'属性を取得、TypeAttr 構造体を返す ----------------------------

Private Function call_GetTypeAttr_(typeInfo_ As IUnknown) As CoTypeAttrLite
    
    
    'パラメータ設定
    
    Dim pAtr_ As LongPtr '受け取る TypeAttr の先頭
    
    
    
    '構造体を取得
    
    Dim resA_
    Dim ppAtr_: ppAtr_ = VarPtr(pAtr_)
    
    Const cGetTypeAttr As LongPtr = 3 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cGetTypeAttr, cStdCall, vbLong, 1, varType(ppAtr_), VarPtr(ppAtr_), resA_
    
    'if res_ =  then err.Raise 0,,""
    
    
    
    '結果を格納
    
    MoveMemory ByVal VarPtr(call_GetTypeAttr_), ByVal pAtr_, LenB(call_GetTypeAttr_)
    
    
    
    
    '構造体開放
    
    Dim resB_
    Dim pAtrRelease_: pAtrRelease_ = pAtr_
    
    Const cReleaseTypeAttr As LongPtr = 19 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cReleaseTypeAttr, cStdCall, vbEmpty, 1, varType(pAtrRelease_), VarPtr(pAtrRelease_), resB_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Function

' - - - - - - - -

'パラメータ設定関数は不要

' ----------------------------------------------------------------





' member pointer 取得 ----------------------------

Public Function Call_AddressOfMember(typeInfo_ As IUnknown, dispId_&) As LongPtr
    
    'パラメータ設定
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForAddressOfMember_ types_, ptrs_, _
        give_dispId_:=dispId_, _
        get_pAddress_:=Call_AddressOfMember
    
    
    
    '取得
    
    Dim res_
    
    Const cAddressOfMember As LongPtr = 15 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cAddressOfMember, cStdCall, vbLong, 3, types_(0), ptrs_(0), res_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Function

' - - - - - - - -

Private Sub setParamsForAddressOfMember_( _
 _
    ByRef ref_types_%(), ByRef ref_ptrs_() As LongPtr, _
 _
    give_dispId_&, _
 _
    ByRef get_pAddress_ As LongPtr _
 _
)
    
    If (Not Not ref_types_) = 0 Then
        
        ReDim ref_types_(3 - 1)
        ReDim ref_ptrs_(3 - 1)
        
        '引数
        Static memid_:      memid_ = give_dispId_
        Static invKind_:    invKind_ = 1&
        Static ppAddr_:     ppAddr_ = VarPtr(get_pAddress_)
        
        ref_types_(0) = vbLong
        ref_ptrs_(0) = VarPtr(memid_)
        
        ref_types_(1) = vbLong
        ref_ptrs_(1) = VarPtr(invKind_)
        
        ref_types_(2) = varType(ppAddr_)
        ref_ptrs_(2) = VarPtr(ppAddr_)
        
    End If
    
    
    memid_ = give_dispId_
    
    ppAddr_ = VarPtr(get_pAddress_)
    
    
End Sub

' ----------------------------------------------------------------





' Invoke ----------------------------

'IDispatch::Invoke() でメソッドを呼び出す。
'　argLength_ がマイナスなら引数リストは既に逆順になっているとする。※ Invoke() は引数配列を逆順で要求する。呼び出し規約の関係なんだろう。
'　※現状では逆順のみ対応、正順の場合はどうしよう vt_ref の配列確保するか…

Public Sub Call_Invoke(object_ As Object, dispId_&, pArgs_ As LongPtr, argLength_&, ByRef out_result_)
    
    Dim argLengthAbs_&: argLengthAbs_ = Abs(argLength_)
    
    
    Dim dp_ As DispParamsStruct
    
    If argLength_ <> 0 Then
        
        '引数情報構造体へセット
        
        dp_.pArgs_ = pArgs_
        
        dp_.argLength_ = argLengthAbs_
        
    End If
    
    
    ' Invoke() のパラメータを設定
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForInvoke_ types_, ptrs_, _
        give_dispId_:=dispId_, _
        give_dispParams_:=dp_, _
        get_result_:=out_result_
    
    
    
    '呼び出し
    
    Dim res_
    
    Const cInvoke As LongPtr = 6 * cSizeOfPointer
    
    DispCallFunc ObjPtr(object_), cInvoke, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), res_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Sub


' - - - - - - - -

Private Sub setParamsForInvoke_( _
 _
    ByRef ref_types_%(), ByRef ref_ptrs_() As LongPtr, _
 _
    give_dispId_&, give_dispParams_ As DispParamsStruct, _
 _
    ByRef get_result_ _
 _
)
    
    If (Not Not ref_types_) = 0 Then
        
        ReDim ref_types_(8 - 1)
        ReDim ref_ptrs_(8 - 1)
        
        Static iid_ As CoGuid       'ＧＵＩＤはゼロ
        
        Const cLocaleSystemDefault  As Long = &H800&
        Const cDispatchMethod       As Long = 1
        Const cDispatchPropertyGet  As Long = 2
        
        '引数
        
        Static dispId_:     dispId_ = give_dispId_
        Static pIid_:       pIid_ = VarPtr(iid_)
        Static locale_:     locale_ = cLocaleSystemDefault
        Static callType_:   callType_ = cDispatchMethod Or cDispatchPropertyGet
        Static pParams_:    pParams_ = VarPtr(give_dispParams_)
        Static pResult_:    pResult_ = VarPtr(get_result_)
        Static pErr1_:      pErr1_ = CLngPtr(0)
        Static pErr2_:      pErr2_ = CLngPtr(0)
        
        
        ref_types_(0) = vbLong
        ref_ptrs_(0) = VarPtr(dispId_)
        
        ref_types_(1) = varType(pIid_)
        ref_ptrs_(1) = VarPtr(pIid_)
        
        ref_types_(2) = vbLong
        ref_ptrs_(2) = VarPtr(locale_)
        
        ref_types_(3) = vbInteger
        ref_ptrs_(3) = VarPtr(callType_)
        
        ref_types_(4) = varType(pParams_)
        ref_ptrs_(4) = VarPtr(pParams_)
        
        ref_types_(5) = varType(pResult_)
        ref_ptrs_(5) = VarPtr(pResult_)
        
        ref_types_(6) = varType(pErr1_)
        ref_ptrs_(6) = VarPtr(pErr1_)
        
        ref_types_(7) = varType(pErr1_)
        ref_ptrs_(7) = VarPtr(pErr1_)
        
    End If
    
    
    dispId_ = give_dispId_
    
    pParams_ = VarPtr(give_dispParams_)
    
    pResult_ = VarPtr(get_result_)
    
    
End Sub


'---------------------------------


Public Function GetAddressOfSafeArray(pSafeArray_ As LongPtr) As LongPtr
    
    Dim sa_ As SafeArrayStruct
    
    MoveMemory sa_, ByVal pSafeArray_, LenB(sa_)
    
    GetAddressOfSafeArray = sa_.pvData
    
End Function



Public Sub FormatCallableParams(length_&)
    
    Dim capacity_&: capacity_ = length_ + length_ \ 2   '必要数の 1.5 倍を確保。
    If capacity_ < 10 Then capacity_ = 10               '最低でも 10 は確保する。
    
    ReDim xCom.CallableParamTypes(capacity_ - 1)
    ReDim xCom.CallableParamPtrs(capacity_ - 1)
    ReDim xCom.CallableParamArgs(capacity_ - 1)
    
    Dim i&
    For i = 0 To capacity_ - 1
        
        xCom.CallableParamArgs(i).varType = vbVariant Or &H4000
        
        xCom.CallableParamPtrs(i) = VarPtr(xCom.CallableParamArgs(i))
        
        xCom.CallableParamTypes(i) = vbVariant Or &H4000
        
    Next
    
End Sub











' EnumVariant ===============================================================================


' EnumVariant を生成する -----------------------------------------------

'ここで生成する EnumVariant は、以下の性質を持つ。
'　・Next で実行する OperationFunc を所持する。
'　・CoTaskMemAlloc/Free() で確保／破棄される。
'　・最低限の機能しかなく、QueryInterface, Clone, Reset, Skip はほぼスタブである。


' EnumVariant を生成する。

Public Function CreateEnumVariant(operationFunc_ As IFunc) As IEnumVARIANT
    
    ' v-table を構築
    
    Static vtable_(7 - 1) As LongPtr ' v-table は全 Enumrator で同じ←staticはインスタンス単位なのでだめみたい
    
    If vtable_(0) = cNullPointer Then setVTable_ vtable_
    
    
    ' EnumVariant を生成
    
    Dim pevar_ As LongPtr
    
    pevar_ = createEnumVariant_(VarPtr(vtable_(0)), operationFunc_)
    
    
    ' IEnumVariant オブジェクトを返す
    
    MoveMemory CreateEnumVariant, pevar_, cSizeOfPointer
    
End Function


'オブジェクト変数はＣＯＭオブジェクト本体への参照アドレスを保持しているにすぎない。
'　ヌルポインタ以外が格納されていれば、参照が破棄された場合等に IUnknown.Release() が走って解放される。
'　逆に言うと、オブジェクト変数が正しく参照カウントを増減できるように調整すれば、MoveMemory() などで差し替えることもできるということ。

Private Function createEnumVariant_(pVTable_ As LongPtr, operationFunc_ As IFunc) As LongPtr
    
    
    ' EnumVariant のメンバを設定する。
    
    Dim evar_ As EnumVariantStruct
    
    evar_.PVtable = pVTable_
    
    evar_.RefCount = 1
    
    Set evar_.OperatorFunc = operationFunc_
    
    
    ' EnumVariant の設定値を確保したヒープメモリにコピーし、IEnumVARIANT の実体を生成する。
    
    Dim pevar_ As LongPtr:  pevar_ = CoTaskMemAlloc(LenB(evar_))
    
    MoveMemory ByVal pevar_, evar_, LenB(evar_)
    
    
    ' evar_ とともにオブジェクトが破棄される前に、クリアしておかなければならない。
    
    Dim nulls_ As EnumVariantStruct
    
    MoveMemory evar_, nulls_, LenB(nulls_)
    
    
    '実体のアドレスを返す。
    
    createEnumVariant_ = pevar_
    
End Function


'ようするに V-table をつくってあげて、任意に確保したヒープメモリをＣＯＭオブジェクトだと思い込ませちゃえばいいって話

Private Function setVTable_(ByRef vtable_() As LongPtr) As LongPtr
    
    vtable_(0) = VBA.CLngPtr(AddressOf xCom.queryInterface_EnumVariant_)
    vtable_(1) = VBA.CLngPtr(AddressOf xCom.addRef_EnumVariant_)
    vtable_(2) = VBA.CLngPtr(AddressOf xCom.release_EnumVariant_)
    vtable_(3) = VBA.CLngPtr(AddressOf xCom.next_EnumVariant_)
    vtable_(4) = VBA.CLngPtr(AddressOf xCom.skip_EnumVariant_)
    vtable_(5) = VBA.CLngPtr(AddressOf xCom.reset_EnumVariant_)
    vtable_(6) = VBA.CLngPtr(AddressOf xCom.clone_EnumVariant_)
    
End Function


' - - - - - - - - - - -

Private Function queryInterface_EnumVariant_(ByRef evar_ As EnumVariantStruct, ByRef riid_ As CoGuid, ByRef pEnumVariant_ As LongPtr) As Long
    
    evar_.RefCount = evar_.RefCount + 1
    
''    Debug.Print GuidToString(riid_)
    
    Select Case True
        
        Case isIUnknown_(riid_)
            
            pEnumVariant_ = VarPtr(evar_)
            
        Case isIEnumVariant_(riid_)
        
            pEnumVariant_ = VarPtr(evar_)
            
        Case Else
            
            queryInterface_EnumVariant_ = cE_NOINTERFACE '取得失敗
            
    End Select
    
''    Debug.Print "query "; VarPtr(evar_); evar_.RefCount
End Function

'条件とかよくわかってないんだけど、variant 型に IEnumVariant をセットしようとしたときにエクセルが落ちることがある。
'いろいろ試してみると、この時 queryInterface() でＧＵＩＤ B196B283-BAB4-101A-B69C-00AA00341D07 を要求しくることに気づいた。
'これは IprovideClassInfo インターフェースというものらしく、唯一？のメソッド GetClassInfo() から ItypeInfo インターフェースを得るためのもの（っぽい）。
'これをちゃんと実装するとたぶん落ちなくなるんだと思うが、かなり大変そうなのでやめとく。代わりに cE_NOINTERFACE を返すようにした。
'とりあえず「型が一致しません」エラーが出て、エクセルが落ちることはなくなった。


Private Function addRef_EnumVariant_(ByRef evar_ As EnumVariantStruct) As Long
    
    evar_.RefCount = evar_.RefCount + 1
    
    addRef_EnumVariant_ = evar_.RefCount
    
''    Debug.Print "addref "; VarPtr(evar_); evar_.RefCount
End Function


Private Function release_EnumVariant_(ByRef evar_ As EnumVariantStruct) As Long
    
    evar_.RefCount = evar_.RefCount - 1
    
    release_EnumVariant_ = evar_.RefCount
    
    
''    Debug.Print "release "; VarPtr(evar_); evar_.RefCount
    If evar_.RefCount = 0 Then
        
        Set evar_.OperatorFunc = Nothing
        
        CoTaskMemFree VarPtr(evar_) '意外にも evar_ 終期化がありそうなのに破棄しても大丈夫みたい
        
    End If
    
End Function


Private Function next_EnumVariant_(ByRef evar_ As EnumVariantStruct, ByVal requestLength_&, ByRef out_Item_, ByVal pCFetch_ As LongPtr) As Long
    
    ' out_item_ は empty なので、オペレータ中での使用は VariantCopy() でよい
    
    next_EnumVariant_ = 1 + evar_.OperatorFunc.xExec02(evar_.OperatorFunc, out_Item_)
    
''    Debug.Print "next "; VarPtr(evar_); evar_.RefCount
End Function



Private Function skip_EnumVariant_(ByRef evar_ As EnumVariantStruct, ByVal cElements As Long) As Long
    Debug.Print "skip "; VarPtr(evar_); evar_.RefCount
End Function


Private Function reset_EnumVariant_(ByRef evar_ As EnumVariantStruct) As Long
    Debug.Print "reset "; VarPtr(evar_); evar_.RefCount
End Function


Private Function clone_EnumVariant_(ByRef evar_ As EnumVariantStruct, ByRef pEnumVariant_ As LongPtr) As Long  'ByRef iev_ As IEnumVARIANT) As Long
    Debug.Print "clone "; VarPtr(evar_); evar_.RefCount
    
    pEnumVariant_ = VarPtr(evar_)
'    Set iev_ = CreateEnumVariant(evar_.OperatorFunc) 'ああ、でもこれデリゲートだったら中のオブジェクトまで複製できないよね…。
    
    clone_EnumVariant_ = 1
    
End Function


'-----------------------------------------------------------------------------------------------



' EnumVariant を使用する -----------------------------------------


'オブジェクトから EnumVariant を取得する。配列はどうしようか…。

Public Function GetEnumVariant(enumerableSourceObject_ As Object) As IEnumVARIANT
    
    If enumerableSourceObject_ Is Nothing Then Exit Function
    
    
    Dim iterator_ 'As IEnumVARIANT
    
    Const cIID_IEnumVARIANT& = -4
    
    xCom.Call_Invoke enumerableSourceObject_, cIID_IEnumVARIANT, 0, 0, iterator_
    
    
    Set GetEnumVariant = iterator_
    
End Function



' EnumVariant の Next() を呼び出す。- - - - - - - - -

'列挙途中なら真を返す。列挙が終了していれば偽を返す。

Public Function CallNext_EnumVariant(src_ As IEnumVARIANT, ByRef out_Item_) As Boolean
    
    'Private Function IEnumVariant::Next( ByVal requestLength_&, ByRef ref_item_, ByVal pCFetch_ as LongPtr ) As Long
    
    Static pItem_
    Static types_(3 - 1) As Integer
    Static pArgs_(3 - 1) As LongPtr
    If types_(0) = 0 Then setArgReferances_ types_, pArgs_, pItem_
    
    pItem_ = VarPtr(out_Item_)
    
    Dim res_&, result_
    
    res_ = DispCallFunc(ObjPtr(src_), 3 * cSizeOfPointer, cStdCall, vbLong, 3, types_(0), pArgs_(0), result_)
    
    CallNext_EnumVariant = (result_ = 0)
    
End Function

Private Sub setArgReferances_(types_%(), pArgs_() As LongPtr, pItem_)
    
    Static requestLength_, pFetch_
    
    requestLength_ = 1&
    pFetch_ = cNullPointer
    
    types_(0) = vbLong
    types_(1) = varType(cNullPointer)
    types_(2) = varType(cNullPointer)
    
    pArgs_(0) = VarPtr(requestLength_)
    pArgs_(1) = VarPtr(pItem_)
    pArgs_(2) = VarPtr(pFetch_)
    
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - -


' ----------------------------------------


