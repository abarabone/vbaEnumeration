VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



#If Win64 Then
    Const cSizeOfVariant& = 2 + 6 + 16  '24
    Const cSizeOfPointer& = 8
    Const cNullPointer^ = 0
#Else
    Const cSizeOfVariant& = 2 + 6 + 8   '16
    Const cSizeOfPointer& = 4
    Const cNullPointer& = 0
#End If


' win API ---------

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst_ As Any, ByRef src_ As Any, ByVal size_&)

Private Declare PtrSafe Function SafeArrayCreate Lib "oleaut32" (ByVal vt_%, ByVal cDims_&, ByRef rgsabound_ As Any) As LongPtr

Private Declare PtrSafe Function SafeArrayCreateVector Lib "oleaut32" (ByVal vt_%, ByVal lLbound_&, ByVal cElements_&) As LongPtr

Private Declare PtrSafe Function SafeArrayDestroy Lib "oleaut32" (ByRef safeArray_ As Any) As Long

Private Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32" (ByRef safeArray_ As Any) As Long

Private Declare PtrSafe Function SafeArrayGetLBound Lib "oleaut32" (ByRef safeArray_ As Any, ByVal nDim_ As Long, ByRef out_lbound_ As Long) As Long

Private Declare PtrSafe Function SafeArrayGetUBound Lib "oleaut32" (ByRef safeArray_ As Any, ByVal nDim_ As Long, ByRef out_ubound_ As Long) As Long

Private Declare PtrSafe Function SafeArrayGetElemsize Lib "oleaut32" (ByRef safeArray_ As Any) As Long

Private Declare PtrSafe Function SafeArrayAccessData Lib "oleaut32" (ByRef safeArray_ As Any, ByRef out_pvData_ As LongPtr) As Long

Private Declare PtrSafe Function SafeArrayUnaccessData Lib "oleaut32" (ByRef safeArray_ As Any) As Long

Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (ByRef dst_ As Any, ByRef src_ As Any) As Long

'--------------------



Const cVarRef& = &H4000
Const cVarArray& = &H2000





'�\���̒�` ============================================================



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




'�ϑ��I Variant �������C�A�E�g

Private Type VariantStructWithCollection
    vt              As Integer
    reserve0        As Integer
    serialLength    As Long 'reserve1    As Integer 'reserve2    As Integer
    '�����𖳎������ʂ����BVariant �̗\�񕔕����g�p����B�ꖕ�̕s���͂����
    pEntity0        As LongPtr
    dimDefs         As Collection 'pEntity1 As LongPtr
    '�������z�u�Ƃ��Ă̓I�u�W�F�N�g�̓|�C���^�[�Ɠ����Ȃ̂ŁANothing �� cNullPointer �Ɠ���
End Type







'�����o�錾 ============================================================





'�z��ێ�

Private arr_ As VariantStructWithCollection ' Variant �^�ϐ��� MoveMemory �ł���悤�ɂ��̃��������C�A�E�g�ŕێ����Ă���B







'�֐���` ============================================================




'�I�u�W�F�N�g�C���X�^���X�̐V�K�쐬

Public Function CNew() As Ary
    
    Set CNew = New Ary
    
End Function





'�z��m�� ------------------------------------------------


'�ꎟ���z�� - - - - -


'�ꎟ���z����m�ہi�������w��j

Public Function Alloc(length_&, Optional varType_ As VbVarType = vbVariant) As Ary
    Set Alloc = Me
    
    Free
    
    
    arr_.vt = varType_ Or cVarArray
    
    arr_.pEntity0 = SafeArrayCreateVector(varType_, 0&, length_)
    
    arr_.serialLength = length_
    
End Function


'�ꎟ���z��m�ہi�擪�Y�����ƒ������w��j

Public Function AllocBound(baseIndex_&, length_&, Optional varType_ As VbVarType = vbVariant) As Ary
    Set AllocBound = Me
    
    Free
    
    
    arr_.vt = varType_ Or cVarArray
    
    arr_.pEntity0 = SafeArrayCreateVector(varType_, baseIndex_, length_)
    
    arr_.serialLength = length_
    
End Function



'�������z�� - - - - -


'�������z��P�v�f�̒�����錾����B�z��m�ۂ���̂ɃR���N�V�������Ƃ��{���]�|���c�B

Public Function DimDef(length_&) As Ary
    Set DimDef = Me
    
    Free
    
    If arr_.dimDefs Is Nothing Then Set arr_.dimDefs = New Collection
    
    
    arr_.dimDefs.Add Array(0&, length_)
    
    
End Function


'�������z��P�v�f�̐擪�Y�����ƒ�����錾����B

Public Function DimDefBound(baseIndex_&, length_&) As Ary
    Set DimDefBound = Me
    
    Free
    
    If arr_.dimDefs Is Nothing Then Set arr_.dimDefs = New Collection
    
    
    arr_.dimDefs.Add Array(baseIndex_, length_)
    
    
End Function


'�������v�f�̐錾�����ƂɁA�������z����m�ۂ���B

Public Function DimAlloc(Optional varType_ As VbVarType = vbVariant) As Ary
    Set DimAlloc = Me
    
    If arr_.dimDefs Is Nothing Then Exit Function
    
    
    
    '�������v�f�̐錾���� safeArray �m�ۗp�\���̂𐶐�����B
    
    Dim safeArrayBounds_() As SafeArrayBound:   ReDim safeArrayBounds_(arr_.dimDefs.Count - 1)
    
    Dim dimDefine_, i&, totalLength_&
    For Each dimDefine_ In arr_.dimDefs
        
        safeArrayBounds_(i).BaseIndex = dimDefine_(0)
        
        safeArrayBounds_(i).Elements = dimDefine_(1)
        
        totalLength_ = totalLength_ + dimDefine_(1)
        
        i = i + 1
    Next
    
    Set arr_.dimDefs = Nothing '�錾���J��
    
    
    ' safeArray ���m�ۂ���B
    
    arr_.vt = varType_ Or cVarArray
    
    arr_.pEntity0 = SafeArrayCreate(varType_, i, safeArrayBounds_(0))
    
    arr_.serialLength = totalLength_
    
End Function


' -----------------------------------------------------------





'�R�s�[�^���[�u ----------------------------


' Variant �Ɋi�[���ꂽ���I�z��� Ary �Ɉړ�����B
'�@�ړ����ɂ� Empty ������B

Public Function MoveFrom(ByRef ref_array_) As Ary
    Set MoveFrom = Me
    
    Free
    
    
    '�ړ���փR�s�[����B
    
    MoveMemory arr_, ref_array_, cSizeOfVariant
    
    If arr_.vt And cVarRef Then Exit Function '�����̌��� Variant �ȊO�ŎQ�ƂɂȂ��Ă���ꍇ�Ȃǁ@�G���[�������ׂ�
    
    'arr_.serialLength =
    
    
    '�ړ����� Empty �ɂ��A�����瑤�ŊJ������Ȃ��悤�ɂ���B
    
    MoveMemory ref_array_, Empty, cSizeOfVariant
    
    
End Function


'�z����R�s�[����B

Public Function CopyFrom(array_) As Ary
    Set CopyFrom = Me
    
    Free
    
    
    VariantCopy arr_, array_
    
    'arr_.serialLength =
    
End Function


' ------------------------------------------------------------






'��� --------------------------------------------------------


'�z��̉��

Public Function Free() As Ary
    Set Free = Me
    
    If arr_.pEntity0 = cNullPointer Then Exit Function
    
    
    '�z��̔j��
    
    SafeArrayDestroy ByVal arr_.pEntity0
    
    
    ' SafeArray �̃N���A
    
    Dim blankArr_ As VariantStructWithCollection
    
    arr_ = blankArr_ '�v�f�錾����������
    
    
End Function

Private Sub Class_Terminate()
    
    Free
    
End Sub

' ---------------------------------------------------




'���擾 -------------------------------------------


'�������Ƃ̗v�f����Ԃ��B�������ȗ������ꍇ�A0 �������w�肵���Ƃ݂Ȃ��B

Public Property Get Length(Optional dimension_&) As Long
    
    SafeArrayGetUBound ByVal arr_.pEntity0, dimension_ + 1, out_ubound_:=Length
    
    Length = Length + 1
    
End Property

Public Property Get BaseIndex(Optional dimension_&) As Long
    
    SafeArrayGetLBound ByVal arr_.pEntity0, dimension_ + 1, out_lbound_:=BaseIndex
    
End Property


'�����̐���Ԃ��B

Public Property Get Dimension() As Long
    
    Dimension = SafeArrayGetDim(ByVal arr_.pEntity0)
    
End Property


'�S�v�f����Ԃ��B

Public Property Get Count() As Long
    
    Count = arr_.serialLength
    
End Property


' --------------------------------------------------------






'�v�f�ւ̃A�N�Z�X --------------------------------------------

'ary_.Item(0) �̂悤�Ɏg�p���邱�Ƃ�z�肵�Ă���B

'�z��̎Q�Ƃ�Ԃ��A�v�f�A�N�Z�X�ɔz��R�s�[�������Ȃ��悤�ɂ��Ă���B
'�@�������AotherArr_ = ary_.Item �̂悤�ɊO���̃o���A���g�ϐ��Ȃǂɓn���Ă͂����Ȃ��B
'�@Ary �N���X���j�������Ɣz��{�̂��j������A���̂��ƎQ�Ƃ�ʂ��ăA�N�Z�X����ƃG�N�Z����������B�ǂ��������񂩂ˁc�B

Public Property Get Item() As Variant
Attribute Item.VB_UserMemId = 0
    
    
    ' SafeArray �Q�� Varriant �\���̂̍\�z
    
    Dim v_ As VariantStructWithCollection
    
    v_.vt = arr_.vt Or cVarRef
    
    v_.pEntity0 = VarPtr(arr_.pEntity0)
    
    
    '�Q�ƍ\���̂�Ԃ��B
    
    MoveMemory Item, v_, cSizeOfVariant
    
    
End Property


'�����𖳎������ʂ��C���f�b�N�X���w�肵�ėv�f���擾����B�擾�ł����ꍇ�͐^�A�ł��Ȃ������ꍇ�͋U��Ԃ��B

Public Function Acquire(serialIndex_&, ByRef out_item_) As Boolean
    
    If serialIndex_ >= arr_.serialLength Then out_item_ = Empty: Exit Function
    
    
    Dim pElements_ As LongPtr
    
    SafeArrayAccessData ByVal arr_.pEntity0, out_pvData_:=pElements_
    
    VariantCopy out_item_, ByVal pElements_ + serialIndex_ * SafeArrayGetElemsize(ByVal arr_.pEntity0)
    
    SafeArrayUnaccessData ByVal arr_.pEntity0
    
    
    Acquire = True
    
End Function

' ------------------------------------------------------------












Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    
'    Set NewEnum = Enumerable(Me).NewEnum �Ɠ���
    
    Dim op_ As New EnumOperatorProcs
    
    Set op_.SourceIterator = Me
    
    Set NewEnum = xCom.CreateEnumVariant(Delegate.CNew(op_, "OpIteratorAry"))
    
End Function




