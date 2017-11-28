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






'�\���̒�` ============================================================



' GUID �i�[�p�\���́i���g�͂�����Ȃ��̂łP�U�o�C�g����΂����j

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





' TypeInfo �I�u�W�F�N�g�̊֐����\����

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




' TypeInfo �I�u�W�F�N�g�̑������\���́i�T�C�Y�Z�k�Łj

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





' Invoke �������

Private Type DispParamsStruct
    pArgs_          As LongPtr
    pArgNames_      As LongPtr
    argLength_      As Long
    argNameLength_  As Long
End Type




' Variant �������C�A�E�g

Private Type VariantStruct
    varType     As Integer
    reserve0    As Integer
    reserve1    As Integer
    reserve2    As Integer
    pEntity0    As LongPtr
    pEntity1    As LongPtr
End Type




'���O EnumVariant �\����

Public Type EnumVariantStruct
    
    PVtable         As LongPtr
    
    OperatorFunc    As IFunc
    
    RefCount        As Long
    
End Type




' SafeArray �������C�A�E�g

Private Type SafeArrayStruct
    
    cDims       As Integer
    
    fFeatures   As Integer
    
    
    cbElements  As Long     '�P�v�f�̃o�C�g�T�C�Y
    
    cLocks      As Long     '���b�N���B������ŃJ�E���g�A�b�v���Ă��܂��ƁA�u�a�`�����瑀��ł��Ȃ��Ȃ�B
    
    pvData      As LongPtr  'PVOID
    
    
    'SafeArrayBound rgsabound[ cDims ];
    
End Type
Private Type SafeArrayBound
    
    Elements    As Long     '�v�f��
    
    BaseIndex   As Long     ' LBound()
    
End Type






' ���p�ϐ� ==========================================================


' CallDispFunc() �p�̔z��B
'�@���W���[�����x���Ő錾���Ȃ��ōςޕ��@��͍��������B
'�@�@�E Static �Ɠ����̌��ʁi��x�s���Έȍ~���ށj
'�@�@�E Redim �ł���
'�@�@�E�֐���ʂ��Ȃ��ōςށi�ł���΁j
'�@�@�Ƃ��������𖞂������@�͂Ȃ��������Ȃ��c�BCoTaskMem() �Ƃ��H

Public CallableParamTypes() As Integer
Public CallableParamPtrs()  As LongPtr
Public CallableParamArgs()  As VariantStruct








'���[�e�B���e�B�֐� =======================================================



'�f�t�h�c�𕶎���ɕϊ�����B

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




' variant �l���󂯎��A����ւ̎Q�� variant �l�𐶐����ĕԂ��B
'�Q�ƌ����j�����ꂽ�ꍇ�A�A�N�Z�X����Ɨ�����̂Œ��ӁB

Public Function ref(value_) As Variant
    
    Dim v_ As VariantStruct
    
    MoveMemory v_, value_, LenB(v_)
    
    
    If (v_.varType And &H4000) = 0 Then
        
        v_.varType = v_.varType Or &H4000
        
        v_.pEntity0 = VarPtr(value_) + (VarPtr(v_.pEntity0) - VarPtr(v_))
        
    End If
    
    
    MoveMemory ref, v_, LenB(v_)
    
End Function





'�󂯎���� variant �l�� Empty �ɂ��āA�R�s�[���� variant ��߂�l�Ƃ��ĕԂ��B
'�@���߂�l��ʂ� variant �ɑ�������ꍇ�́A�z��̃R�s�[�͍s���Ȃ��i�炵���A���؂��������܂����킩��񂩂����j�B

Public Function Move(value_) As Variant
    
    MoveMemory Move, value_, cSizeOfVariant
    
    MoveMemory value_, Empty, cSizeOfVariant '������ɂ��Ȃ��ƊJ�����Q�d�ɍs���ė�����B
    
End Function







'������

Public Sub CallFunc2(pObj_ As LongPtr, pFunc_ As LongPtr, ByRef args_, Optional ByRef out_result_, Optional isVbaClass_ As Boolean)
    
    Dim argLength_&:    argLength_ = UBound(args_) + 1
    
    Dim argResLength_&: argResLength_ = argLength_ + (Not IsMissing(out_result_)) * isVbaClass_
    
    
    
    '�����p�z��̏�����
    
    Static types_%()
    Static pArgs_() As LongPtr
    
    Select Case True
        
        Case (Not types_) = -1 '��������
            
            Dim minLen_&: minLen_ = IIf(argResLength_ > 7, argResLength_, 7) '�Œ� 10 �͊m�ۂ����B
            
            initArgArray_ minLen_, types_, pArgs_
            
        Case argResLength_ > UBound(types_) + 1
            
            initArgArray_ argResLength_, types_, pArgs_
            
    End Select
    
    
    '�����̃Z�b�g
    
    Dim i&
    For i = 0 To argLength_ - 1
        
        pArgs_(i) = VarPtr(args_(i))
        
    Next
    
    
    
    '�֐����R�[��
    
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
    
    
    
    '�����p�z��̏�����
    
    Static types_%()
    Static pArgs_() As LongPtr
    
    Select Case True
        
        Case (Not types_) = -1 '��������
            
            Dim minLen_&: minLen_ = IIf(argResLength_ > 7, argResLength_, 7) '�Œ� 10 �͊m�ۂ����B
            
            initArgArray_ minLen_, types_, pArgs_
            
        Case argResLength_ > UBound(types_) + 1
            
            initArgArray_ argResLength_, types_, pArgs_
            
    End Select
    
    
    '�����̃Z�b�g
    
    Dim i&
    For i = 0 To argLength_ - 1
        
        pArgs_(i) = pArgTop_ + i * cSizeOfVariant
        
    Next
    
    
    
    '�֐����R�[��
    
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

'�����p�z�������������B

Private Sub initArgArray_(length_&, ByRef types_%(), ByRef pArgs_() As LongPtr)
    
    Dim capacity_&: capacity_ = length_ + length_ \ 2 - 1 '�K�v���� 1.5 �{���m�ہB
    
    ReDim types_(capacity_)
    ReDim pArgs_(capacity_)
    
    Dim i&
    For i = 0 To capacity_
        
        types_(i) = vbVariant Or &H4000
        
    Next
    
End Sub

'�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|









' dispId ���烁�\�b�h���Ăяo���B

Public Function InvokeMethod(object_ As Object, dispId_&, ParamArray args_()) As Variant
    
    Call_Invoke object_, dispId_, VarPtr(args_(0)), UBound(args_) + 1, out_result_:=InvokeMethod
    
End Function










'���\�b�h���O���珘�����擾����

Public Function GetOrderByName(object_ As Object, methodName_$) As Long
    
    Dim typeInfo_ As IUnknown
    Set typeInfo_ = Call_GetTypeInfo(object_)
    
    Dim dispId_&:   dispId_ = Call_GetIDsOfNames(object_, methodName_)
    
    
    Dim i&
    For i = 0 To call_GetTypeAttr_(typeInfo_).cFuncs - 1
        
        ' funcDesc �\���̂��擾���� dispId ����v���郁�\�b�h��T���Avtable ��̏��Ԃ��擾����B
        
        Dim desc_ As CoFuncDesc
        desc_ = call_GetFuncDesc_(typeInfo_, i)
        
        If desc_.memid = dispId_ Then
            
            GetOrderByName = desc_.oVft
            
            Exit Function
            
        End If
        
    Next
    
    GetOrderByName = -1
    
End Function







'�I�u�W�F�N�g�����o���\�b�h�����X�g�̎擾

Public Function MakeVtableList(object_ As Object) As Collection
    
    Dim typeInfo_ As IUnknown
    Set typeInfo_ = Call_GetTypeInfo(object_)
    
    
    Dim c_ As New Collection
    
    
    Dim i&
    For i = 0 To call_GetTypeAttr_(typeInfo_).cFuncs - 1
        
        
        ' funcDesc �\���̂��擾���āAmemid �� vtable ��̏��Ԃ��擾����B
        
        Dim desc_ As CoFuncDesc
        desc_ = call_GetFuncDesc_(typeInfo_, i)
        
        
        ' dispId_ ���烁�\�b�h�̖��O���擾����B
        
        Dim thisName_$
        thisName_ = call_GetNames(typeInfo_, desc_.memid)
        
        
        c_.Add desc_.oVft, thisName_
        
    Next
    
    
    Set MakeVtableList = c_
    
End Function








'���\�b�h�A�h���X���擾����B

Public Function GetAddress(object_ As Object, methodName_$) As LongPtr
    
    Dim typeInfo_ As IUnknown
    Set typeInfo_ = Call_GetTypeInfo(object_)
    
    
    Dim dispId_&
    dispId_ = Call_GetIDsOfNames(object_, methodName_)
    
    
    GetAddress = Call_AddressOfMember(typeInfo_, dispId_)
    
End Function



' v-table ���烁�\�b�h�A�h���X���擾����B

Public Function GetAddressFromVtable(object_ As Object, methodName_$) As LongPtr
'    Dim ti_: Set ti_ = call_GetTypeInfo_(object_)
    Dim order_&:    order_ = xCom.GetOrderByName(object_, methodName_)
    
    
    Dim pVTable_ As LongPtr
    
    MoveMemory ByVal VarPtr(pVTable_), ByVal ObjPtr(object_), LenB(pVTable_)
    
    MoveMemory ByVal VarPtr(GetAddressFromVtable), ByVal pVTable_ + order_, LenB(GetAddressFromVtable)
    
'    Dim a_(64) As LongPtr
'    MoveMemory ByVal VarPtr(a_(0)), ByVal pVTable_, 32 * 4
    
End Function











'�C���^�[�t�F�[�X���\�b�h�R�[�� ===============================================================================




' TypeInfo �擾 ----------------------------

Public Function Call_GetTypeInfo(object_ As Object) As IUnknown
    
    
    '�p�����[�^�ݒ�
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForGetTypeInfo_ types_, ptrs_, _
        get_typeInfo_:=Call_GetTypeInfo
    
    
    
    '�擾
    
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
        
        '����
        Static zero_:   zero_ = 0&                      ' 0 �ł� Integer �ɂȂ�̂Ń_��
        Static locale_: locale_ = &H800&                ' LOCALE_SYSTEM_DEFAULT
        Static ppTi_:   ppTi_ = VarPtr(get_typeInfo_)   ' IUnknown �^�̂��|�C���^�l�����������Ď擾����B
        
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





' dispid �擾 ----------------------------

Public Function Call_GetIDsOfNames(object_ As Object, fname_$) As Long
    
    
    '�p�����[�^�ݒ�
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForGetDispId_ types_, ptrs_, _
        give_fname_:=fname_, _
        get_dispId_:=Call_GetIDsOfNames
    
    
    
    '�擾
    
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
        
        Static iid_ As CoGuid       '�f�t�h�c�̓[��
        Static pName_ As LongPtr    '�֐���
        
        '����
        Static pIid_:   pIid_ = VarPtr(iid_)
        Static ppName_: ppName_ = VarPtr(pName_)
        Static count_:  count_ = 1&                 ' 1 �ł� Integer �ƂȂ�o�O��A���̂ւ�V�r�A
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






'���\�b�h�����擾�AFuncDesc �\���̂�Ԃ� ----------------------------

Private Function call_GetFuncDesc_(typeInfo_ As IUnknown, funcIndex_&) As CoFuncDesc
    
    
    '�p�����[�^�ݒ�
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    Dim pDesc_ As LongPtr '���� FuncDesc �̐擪
    
    setParamsForGetFuncDesc_ types_, ptrs_, _
        give_funcIndex_:=funcIndex_, _
        get_pDesc_:=pDesc_
    
    
    
    
    ' FuncDesc �\���̂̎擾
    
    Dim resA_
    
    Const cGetFuncDesc As LongPtr = 5 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cGetFuncDesc, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), resA_
    
    'if res_ =  then err.Raise 0,,""
    
    
    
    '�\���̂��R�s�[����
    
    MoveMemory ByVal VarPtr(call_GetFuncDesc_), ByVal pDesc_, LenB(call_GetFuncDesc_)
    
    
    
    ' FuncDesc �\���̂̊J��
    
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
        
        '����
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








'���O�̎擾 ----------------------------

Public Function call_GetNames(typeInfo_ As IUnknown, dispId_&) As String
    
    
    '�p�����[�^�ݒ�
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    Dim pName_ As LongPtr '�󂯎�閼�O������̐擪�i�Ȃ�ŊJ�����Ȃ��Ă����񂾂낤�H�j
    
    setParamsForGetName_ types_, ptrs_, _
        give_dispId_:=dispId_, _
        get_pName_:=pName_
    
    
    
    '�擾
    
    Dim res_
    
    Const cGetNames As LongPtr = 7 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cGetNames, cStdCall, vbLong, UBound(types_) + 1, types_(0), ptrs_(0), res_ ' get names
    
    'if res_ =  then err.Raise 0,,""
    
    
    
    '���ʂ��i�[����B
    
    MoveMemory ByVal VarPtr(call_GetNames), ByVal VarPtr(pName_), LenB(pName_) ' pName_ �� bstr
    
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
        
        '�Q�Ɛ�
        Static pRtnLength_& '�擾�����Ԃ��Ă��邪�A�g�p���Ȃ��B
        
        '����
        Static dispId_:     dispId_ = give_dispId_
        Static maxLength_:  maxLength_ = 1&                 ' 1 �ł� Integer �ƂȂ�o�O��A���̂ւ�V�r�A
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




'�������擾�ATypeAttr �\���̂�Ԃ� ----------------------------

Private Function call_GetTypeAttr_(typeInfo_ As IUnknown) As CoTypeAttrLite
    
    
    '�p�����[�^�ݒ�
    
    Dim pAtr_ As LongPtr '�󂯎�� TypeAttr �̐擪
    
    
    
    '�\���̂��擾
    
    Dim resA_
    Dim ppAtr_: ppAtr_ = VarPtr(pAtr_)
    
    Const cGetTypeAttr As LongPtr = 3 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cGetTypeAttr, cStdCall, vbLong, 1, varType(ppAtr_), VarPtr(ppAtr_), resA_
    
    'if res_ =  then err.Raise 0,,""
    
    
    
    '���ʂ��i�[
    
    MoveMemory ByVal VarPtr(call_GetTypeAttr_), ByVal pAtr_, LenB(call_GetTypeAttr_)
    
    
    
    
    '�\���̊J��
    
    Dim resB_
    Dim pAtrRelease_: pAtrRelease_ = pAtr_
    
    Const cReleaseTypeAttr As LongPtr = 19 * cSizeOfPointer
    
    DispCallFunc ObjPtr(typeInfo_), cReleaseTypeAttr, cStdCall, vbEmpty, 1, varType(pAtrRelease_), VarPtr(pAtrRelease_), resB_
    
    'if res_ =  then err.Raise 0,,""
    
    
End Function

' - - - - - - - -

'�p�����[�^�ݒ�֐��͕s�v

' ----------------------------------------------------------------





' member pointer �擾 ----------------------------

Public Function Call_AddressOfMember(typeInfo_ As IUnknown, dispId_&) As LongPtr
    
    '�p�����[�^�ݒ�
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForAddressOfMember_ types_, ptrs_, _
        give_dispId_:=dispId_, _
        get_pAddress_:=Call_AddressOfMember
    
    
    
    '�擾
    
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
        
        '����
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

'IDispatch::Invoke() �Ń��\�b�h���Ăяo���B
'�@argLength_ ���}�C�i�X�Ȃ�������X�g�͊��ɋt���ɂȂ��Ă���Ƃ���B�� Invoke() �͈����z����t���ŗv������B�Ăяo���K��̊֌W�Ȃ񂾂낤�B
'�@������ł͋t���̂ݑΉ��A�����̏ꍇ�͂ǂ����悤 vt_ref �̔z��m�ۂ��邩�c

Public Sub Call_Invoke(object_ As Object, dispId_&, pArgs_ As LongPtr, argLength_&, ByRef out_result_)
    
    Dim argLengthAbs_&: argLengthAbs_ = Abs(argLength_)
    
    
    Dim dp_ As DispParamsStruct
    
    If argLength_ <> 0 Then
        
        '�������\���̂փZ�b�g
        
        dp_.pArgs_ = pArgs_
        
        dp_.argLength_ = argLengthAbs_
        
    End If
    
    
    ' Invoke() �̃p�����[�^��ݒ�
    
    Static types_%()
    Static ptrs_() As LongPtr
    
    setParamsForInvoke_ types_, ptrs_, _
        give_dispId_:=dispId_, _
        give_dispParams_:=dp_, _
        get_result_:=out_result_
    
    
    
    '�Ăяo��
    
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
        
        Static iid_ As CoGuid       '�f�t�h�c�̓[��
        
        Const cLocaleSystemDefault  As Long = &H800&
        Const cDispatchMethod       As Long = 1
        Const cDispatchPropertyGet  As Long = 2
        
        '����
        
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
    
    Dim capacity_&: capacity_ = length_ + length_ \ 2   '�K�v���� 1.5 �{���m�ہB
    If capacity_ < 10 Then capacity_ = 10               '�Œ�ł� 10 �͊m�ۂ���B
    
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


' EnumVariant �𐶐����� -----------------------------------------------

'�����Ő������� EnumVariant �́A�ȉ��̐��������B
'�@�ENext �Ŏ��s���� OperationFunc ����������B
'�@�ECoTaskMemAlloc/Free() �Ŋm�ہ^�j�������B
'�@�E�Œ���̋@�\�����Ȃ��AQueryInterface, Clone, Reset, Skip �͂قڃX�^�u�ł���B


' EnumVariant �𐶐�����B

Public Function CreateEnumVariant(operationFunc_ As IFunc) As IEnumVARIANT
    
    ' v-table ���\�z
    
    Static vtable_(7 - 1) As LongPtr ' v-table �͑S Enumrator �œ�����static�̓C���X�^���X�P�ʂȂ̂ł��߂݂���
    
    If vtable_(0) = cNullPointer Then setVTable_ vtable_
    
    
    ' EnumVariant �𐶐�
    
    Dim pevar_ As LongPtr
    
    pevar_ = createEnumVariant_(VarPtr(vtable_(0)), operationFunc_)
    
    
    ' IEnumVariant �I�u�W�F�N�g��Ԃ�
    
    MoveMemory CreateEnumVariant, pevar_, cSizeOfPointer
    
End Function


'�I�u�W�F�N�g�ϐ��͂b�n�l�I�u�W�F�N�g�{�̂ւ̎Q�ƃA�h���X��ێ����Ă���ɂ����Ȃ��B
'�@�k���|�C���^�ȊO���i�[����Ă���΁A�Q�Ƃ��j�����ꂽ�ꍇ���� IUnknown.Release() �������ĉ�������B
'�@�t�Ɍ����ƁA�I�u�W�F�N�g�ϐ����������Q�ƃJ�E���g�𑝌��ł���悤�ɒ�������΁AMoveMemory() �Ȃǂō����ւ��邱�Ƃ��ł���Ƃ������ƁB

Private Function createEnumVariant_(pVTable_ As LongPtr, operationFunc_ As IFunc) As LongPtr
    
    
    ' EnumVariant �̃����o��ݒ肷��B
    
    Dim evar_ As EnumVariantStruct
    
    evar_.PVtable = pVTable_
    
    evar_.RefCount = 1
    
    Set evar_.OperatorFunc = operationFunc_
    
    
    ' EnumVariant �̐ݒ�l���m�ۂ����q�[�v�������ɃR�s�[���AIEnumVARIANT �̎��̂𐶐�����B
    
    Dim pevar_ As LongPtr:  pevar_ = CoTaskMemAlloc(LenB(evar_))
    
    MoveMemory ByVal pevar_, evar_, LenB(evar_)
    
    
    ' evar_ �ƂƂ��ɃI�u�W�F�N�g���j�������O�ɁA�N���A���Ă����Ȃ���΂Ȃ�Ȃ��B
    
    Dim nulls_ As EnumVariantStruct
    
    MoveMemory evar_, nulls_, LenB(nulls_)
    
    
    '���̂̃A�h���X��Ԃ��B
    
    createEnumVariant_ = pevar_
    
End Function


'�悤����� V-table �������Ă����āA�C�ӂɊm�ۂ����q�[�v���������b�n�l�I�u�W�F�N�g���Ǝv�����܂����Ⴆ�΂������Ęb

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
            
            queryInterface_EnumVariant_ = cE_NOINTERFACE '�擾���s
            
    End Select
    
''    Debug.Print "query "; VarPtr(evar_); evar_.RefCount
End Function

'�����Ƃ��悭�킩���ĂȂ��񂾂��ǁAvariant �^�� IEnumVariant ���Z�b�g���悤�Ƃ����Ƃ��ɃG�N�Z���������邱�Ƃ�����B
'���낢�뎎���Ă݂�ƁA���̎� queryInterface() �łf�t�h�c B196B283-BAB4-101A-B69C-00AA00341D07 ��v�������邱�ƂɋC�Â����B
'����� IprovideClassInfo �C���^�[�t�F�[�X�Ƃ������̂炵���A�B��H�̃��\�b�h GetClassInfo() ���� ItypeInfo �C���^�[�t�F�[�X�𓾂邽�߂̂��́i���ۂ��j�B
'����������Ǝ�������Ƃ��Ԃ񗎂��Ȃ��Ȃ�񂾂Ǝv�����A���Ȃ��ς����Ȃ̂ł�߂Ƃ��B����� cE_NOINTERFACE ��Ԃ��悤�ɂ����B
'�Ƃ肠�����u�^����v���܂���v�G���[���o�āA�G�N�Z���������邱�Ƃ͂Ȃ��Ȃ����B


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
        
        CoTaskMemFree VarPtr(evar_) '�ӊO�ɂ� evar_ �I���������肻���Ȃ̂ɔj�����Ă����v�݂���
        
    End If
    
End Function


Private Function next_EnumVariant_(ByRef evar_ As EnumVariantStruct, ByVal requestLength_&, ByRef out_Item_, ByVal pCFetch_ As LongPtr) As Long
    
    ' out_item_ �� empty �Ȃ̂ŁA�I�y���[�^���ł̎g�p�� VariantCopy() �ł悢
    
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
'    Set iev_ = CreateEnumVariant(evar_.OperatorFunc) '�����A�ł�����f���Q�[�g�������璆�̃I�u�W�F�N�g�܂ŕ����ł��Ȃ���ˁc�B
    
    clone_EnumVariant_ = 1
    
End Function


'-----------------------------------------------------------------------------------------------



' EnumVariant ���g�p���� -----------------------------------------


'�I�u�W�F�N�g���� EnumVariant ���擾����B�z��͂ǂ����悤���c�B

Public Function GetEnumVariant(enumerableSourceObject_ As Object) As IEnumVARIANT
    
    If enumerableSourceObject_ Is Nothing Then Exit Function
    
    
    Dim iterator_ 'As IEnumVARIANT
    
    Const cIID_IEnumVARIANT& = -4
    
    xCom.Call_Invoke enumerableSourceObject_, cIID_IEnumVARIANT, 0, 0, iterator_
    
    
    Set GetEnumVariant = iterator_
    
End Function



' EnumVariant �� Next() ���Ăяo���B- - - - - - - - -

'�񋓓r���Ȃ�^��Ԃ��B�񋓂��I�����Ă���΋U��Ԃ��B

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


