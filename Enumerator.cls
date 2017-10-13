VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit





'�񋓉\�I�u�W�F�N�g

'�@��{�I�ɂ́A�k�h�m�p�I�ȃI�y���[�^������s�����߂̃I�u�W�F�N�g�ł���B

'�@�\�[�X�́A�񋓉\�I�u�W�F�N�g�`�F�C��������ȊO�̗񋓉\�ȏW���i�z���R���N�V�����Ȃǁj�łȂ���΂Ȃ�Ȃ��B

'�@�񋓉\�ȏW�����\�[�X�Ƃ��ď�������ꍇ�A������C�e���[�V��������K�v��������B
'�@�@���̂��߂ɕK���񋓉\�ȏW����񋓎q�i�C�e���[�V�����@�\�֐������񋓉\�I�u�W�F�N�g�j�ł���ށB






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







'�\���̒�` ============================================================




' Variant �������C�A�E�g

Private Type EnumVariantStruct
    
    PVtable     As LongPtr
    Operator    As IFunc
    
    RefCount    As Long
    
End Type









'�����o�錾 ============================================================




Private enumerableSource_   As IUnknown     '�Q�ƌ��ƂȂ�񋓉\�ȏW���B�\�[�X�BIEnumVariant ��Ԃ�����́iEnumerator �܂ށj

Private operation_          As Delegate









'�֐���` ============================================================





' For Each ���� EnumVariant ��Ԃ� ------------------------------------------


' enumerableSource_ ������`�̏ꍇ�́ANothing ���Ԃ�B

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    
    Select Case True
        
        Case TypeOf enumerableSource_ Is Ary
            
            Set NewEnum = xCom.CreateEnumVariant(instantiateOperator_(operation_, enumerableSource_))
            
            
        Case TypeOf enumerableSource_ Is Enumerator
            
            Set NewEnum = xCom.CreateEnumVariant(instantiateOperator_(operation_, asObject_(enumerableSource_).NewEnum))
            
            
        Case TypeOf enumerableSource_ Is Object
            
            Dim ienv_
            
            Const cIID_IEnumVARIANT& = -4
            
            xCom.Call_Invoke enumerableSource_, cIID_IEnumVARIANT, 0, 0, ienv_
            
            Set NewEnum = ienv_
            
            
        Case TypeOf enumerableSource_ Is IEnumVARIANT
            
            '�\�[�X�� EnumVariant �̏ꍇ�́A�N���[���ł����悤�Ǝv�����A�Ƃ肠�����͂��̂܂܂������B
            
            Set NewEnum = enumerableSource_
            
            
    End Select
    
End Function


' ---------------------------------------------------------------------------









'�񋓃\�[�X�ƃI�y���[�V�����f���Q�[�g���Z�b�g���� ------------------------------------------



'�\�[�X�I�u�W�F�N�g�� EnumVariant ��Ԃ��Ȃ��ꍇ�A�f���Q�[�g�ŃC�e���[�V�������L�q���A�\�[�X�I�u�W�F�N�g�ƂƂ��ɃZ�b�g����B

Public Function SetIterator(enumerableSourceObject_ As Object, operationFunction_ As Delegate) As Enumerator
    Set SetIterator = Me
    
    Set enumerableSource_ = enumerableSourceObject_
    
    Set operation_ = operationFunction_
    
End Function



'�\�[�X�ƂȂ� Enumrator �Ƃk�h�m�p�I�I�y���[�^�f���Q�[�g���Z�b�g����B

Public Function SetOperator(sourceEnumerator_ As Enumerator, operationFunction_ As Delegate) As Enumerator
    Set SetOperator = Me
    
    Set enumerableSource_ = sourceEnumerator_
    
    Set operation_ = operationFunction_
    
End Function



'�\�[�X�ƂȂ�I�u�W�F�N�g���Z�b�g����BEnumVariant ���n����B

Public Function SetSource(enumerableSourceObject_ As IUnknown) As Enumerator
    Set SetSource = Me
    
    Set enumerableSource_ = enumerableSourceObject_
    
End Function



' ---------------------------------------------------------------------------







'�k�h�m�p�I�I�y���[�^�[ -----------------------------------------------------------

'�I�y���[�^�[�̃f���Q�[�g�͗񋓎��� Instantiate() ����A�񋓎��̏�Ԃ��ێ�����B

'�@�I�y���[�^�I�u�W�F�N�g�́A�ȉ��̃����o�����B
'�@�@�E SourceIterator  ... EnumVariant �܂��� �񋓃\�[�X���̂���
'�@�@�E Expression      ... �I�y���[�^�Ŏg�p���鏈���̃f���Q�[�g�B
'�@�@�E Variable        ... �񋓒��ɕϓ����A�񋓂��Ƃɏ����������l�B
'
'�I�y���[�^�� Enumrator �𐶐�����ہA�����̒l���Z�b�g�E����������B





Public Function qSelect(selector_) As Enumerator
    Set qSelect = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    Set op_.Expression = toFunc_(selector_)
    
    
    qSelect.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpSelect")
    
End Function




Public Function qWhere(predicate_) As Enumerator
    Set qWhere = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    Set op_.Expression = toFunc_(predicate_)
    
    
    qWhere.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpWhere")
    
End Function




Public Function qTake(count_&) As Enumerator
    Set qTake = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    op_.Variable = count_
    
    
    qTake.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpTake")
    
End Function




Public Function qSkip(count_&) As Enumerator
    Set qSkip = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    op_.Variable = count_
    
    
    qSkip.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpSkip")
    
End Function




Public Function qSpan(count_&) As Enumerator
    Set qSpan = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    op_.Variable = count_
    
    
    qSpan.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpSpan")
    
End Function




Public Function qSelectMany(Optional collectionSelector_, Optional resultSelector_) As Enumerator
    Set qSelectMany = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    op_.Expression = Array(toFunc_(collectionSelector_), toFunc_(resultSelector_))
    
    
    Dim innerIterator_ As IEnumVARIANT '�����C�e���[�^�ϐ��̏�����
    
    Set op_.Variable = innerIterator_
    
    
    qSelectMany.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpSelectMany")
    
End Function





Public Function qGroupBy(Optional keySelector_, Optional elementSelector_, Optional resultSelector_) As Enumerator
    Set qGroupBy = New Enumerator
    
    
    Dim op_ As New EnumOperatorProcs
    
    op_.Expression = Array(toFunc_(keySelector_), toFunc_(elementSelector_), toFunc_(resultSelector_))
    
    
    Dim innerIterator_      As IEnumVARIANT '�����C�e���[�^�ϐ��̏�����
    Dim groupDictionary_    As Dictionary
    Dim key_
    
    op_.Variable = Array(groupDictionary_, key_, innerIterator_)
    
    
    qGroupBy.SetOperator Me, Delegate.CNew.SetMethod(op_, "OpGroupBy")
    
End Function






Private Function xFlatAll(Optional resultSelector_) As Enumerator
    
    Set xFlatAll = New Enumerator
    
    Dim innerIterator_ As Enumerator: Set innerIterator_ = Nothing
    
    
End Function





'�������z��̎������P�����A���ʎ����̗񋓂ɕϊ�����B�o���邩�H�H

'�@[ i, j, k, ... ] -> [ i ][ j, k, ... ] �Ƃ��Ĉ����Ai �ŉ񂷁B�H�H ���������z��͊K�w�^�񋓂Ƃ݂Ȃ��āA���R���ł��ł���������

Private Function qDownDimension() As Enumerator
    
    
    
End Function

'�Ӗ��Ȃ������H
'���������S�Ă̗v�f�̈Ӗ��������Ȃ�΁i�P�Ȃ���W�I�ȈӖ��ł����Ȃ��Ȃ�j�A�t���b�g�ȗ񋓂ł悢�B
'�t�ɗႦ�΂b�r�u�̃e�[�u����G�N�Z���̕\�݂����Ȃ��̂��Ƃ�����A����͍\���I�Ɉ����d�g�݂����ׂ����낤�B
'�@�����ꂪ�߂�ǂ������ꍇ�A�Ƃ������v�����邩�H





'�J�E���g���Ƃɂ܂Ƃ߂āA�O���C�e���[�^���쐬����B���ɗ����̂͊O���C�e���[�^�ŁA�G���g���[�͓����C�e���[�^�ł���B�ł���H�H

Private Function qLot(count_&) As Enumerator
    
    
    
End Function






'���ׂĂ̗v�f�𕽒R�����A�e�v�f���܂߂Ďˉe����B

Private Function qExprole(Optional relaySelector_, Optional terminalSelector_) As Enumerator
    
    
    
End Function










' Expression �f���Q�[�g�ݒ胁�\�b�h - - - - - - - - -

'�I�y���[�^���Z�b�g������ŁA�Z���N�^�Ȃǂ�ݒ�ł���B
'�@�I�y���[�^�ɑΉ����Ă��Ȃ�����ݒ肷��ƁA�G���[�Œʒm����B


Public Function SetCollectionSelector(collectionSelector_) As Enumerator
    Set SetCollectionSelector = Me
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpSelectMany"):  i = 0
         
        Case Else: Exit Function '�Ή����Ă��Ȃ��G���[
    
    End Select
    
    operation_.Target.SetExpression toFunc_(collectionSelector_), i
    
End Function


Public Function SetResultSelector(resutlSelector_) As Enumerator
    Set SetResultSelector = Me
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpSelectMany"):  i = 1
        Case operation_.HasMethod("OpGroupBy"):     i = 2
         
        Case Else: Exit Function '�Ή����Ă��Ȃ��G���[
    
    End Select
    
    operation_.Target.SetExpression toFunc_(resutlSelector_), i
    
End Function


Public Function SetKeySelector(keySelector_) As Enumerator
    Set SetKeySelector = Me
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpGroupBy"):  i = 0
         
        Case Else: Exit Function '�Ή����Ă��Ȃ��G���[
    
    End Select
    
    operation_.Target.SetExpression toFunc_(keySelector_), i
    
End Function


Public Function SetElementSelector(elementSelector_) As Enumerator
    Set SetElementSelector = Me
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpGroupBy"):  i = 1
         
        Case Else: Exit Function '�Ή����Ă��Ȃ��G���[
        
    End Select
    
    operation_.Target.SetExpression toFunc_(elementSelector_), i
    
End Function

' - - - - - - - - - - - - -



' ---------------------------------------------------------------------------------







'�ϊ� -----------------------------------------------------------------------------



'�R���N�V������

'�@�����R���N�V�����ł��A�R�s�[����B

Public Function ToCollection() As Collection
    
    Set ToCollection = New Collection
    
    Dim i
    For Each i In Me
        
        ToCollection.Add i
        
    Next
    
End Function

Public Function ToCollectionAsEnumerable() As Enumerator
    
    Set ToCollectionAsEnumerable = Enumerable(ToCollection)
    
End Function






'�z��
'�@�̈�m�ۂɌ��̎擾���K�v�Ȃ��߁A��U�S�v�f���R���N�V���������Ă���z��̃R�s�[���Ƃ��Ă���B

Public Function ToAry(Optional baseIndex_&) As Ary
    
    Dim src_
    
    If HasCountable _
        Then Set src_ = enumerableSource_ _
        Else Set src_ = ToCollection
    
    Set ToAry = Ary.CNew.AllocBound(baseIndex_, src_.Count)
    
    Dim i&, element_
    For Each element_ In src_
        
        VariantCopy ToAry.Item(i), element_
        
        i = i + 1
    Next
    
End Function

Public Function ToAryAsEnumerable(Optional baseIndex_&) As Enumerator
    
    Set ToAryAsEnumerable = Enumerable(ToAry(baseIndex_))
    
End Function





Private Function ToJaggedArray() As Variant
    
    
    
End Function






'�O���[�s���O�������\�z����B
'�@qGroupBy() �ŗ񋓂������̂Ɠ����B������͗񋓂����Ɏ�����Ԃ��p�^�[���B

' keySelector       �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@element : variant =>> key : variant

' elementSelector   �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@element : variant =>> element_result : variant

Public Function ToLookUp(Optional keySelector_, Optional elementSelector_) As Dictionary
    
    Set ToLookUp = makeGroupDictionary_(Me.NewEnum, toFunc_(keySelector_), toFunc_(elementSelector_))
    
End Function


Private Function makeGroupDictionary_(iterator_ As IEnumVARIANT, keySelector_ As IFunc, elmSelector_ As IFunc) As Dictionary
    
    Dim groups_ As New Dictionary
    
    
    Dim item_
    
    Do While xCom.CallNext_EnumVariant(iterator_, out_item_:=item_)
        
        
        Dim key_
        
        If keySelector_ Is Nothing _
            Then VariantCopy key_, item_ _
            Else VariantCopy key_, keySelector_.xExec01(item_)
        
        
        Dim element_
        
        If elmSelector_ Is Nothing _
            Then VariantCopy element_, item_ _
            Else VariantCopy element_, elmSelector_.xExec01(item_)
        
        
        
        '�L�[�������I�u�W�F�N�g�Ȃ�A�S�v�f�𕶎��񉻂��ăV���A���C�Y����B
        '�@�����I�u�W�F�N�g�́A
        
        Dim stringKey_$
        
        stringKey_ = Grouping.SeriarizeIfAnonymous(key_)
        
        
        
        '�v�f���O���[�v�R���N�V�����ɒǉ�����B
        
        getGrouping_(groups_, key_, stringKey_).Elements.Add element_
        
        
    Loop
    
    
    Set makeGroupDictionary_ = groups_
    
End Function


'�v�f���O���[�v�R���N�V�����ɒǉ�����
'�@�������� Grouping �����o���B�o�^����Ă��Ȃ���ΐV�K Grouping ��o�^����B

Private Function getGrouping_(groups_ As Dictionary, key_, stringKey_$) As Grouping
    
    
    ' key_ �������ɓo�^����Ă���΂����Ԃ��B
    
    On Error Resume Next
    
    If stringKey_ = "" _
        Then Set getGrouping_ = groups_(key_) _
        Else Set getGrouping_ = groups_(stringKey_)
    
    If Not getGrouping_ Is Nothing Then Exit Function
    
    
    '�o�^����Ă��Ȃ���ΐV�K�쐬�A�o�^�A�����ĕԂ��B
    
    Set getGrouping_ = New Grouping
    
    If IsObject(key_) _
        Then Set getGrouping_.Key = key_ _
        Else getGrouping_.Key = key_
    
    
    If stringKey_ = "" _
        Then Set groups_(key_) = getGrouping_ _
        Else Set groups_(stringKey_) = getGrouping_
    
End Function







'�񋓂��玫�����\�z����B

' keySelector       �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@element : variant =>> key : variant

' elementSelector   �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@element : variant =>> element_result : variant

Public Function ToDictionary(Optional keySelector_, Optional elementSelector_) As Dictionary
    
    Set ToDictionary = makeDictionary_(Me.NewEnum, toFunc_(keySelector_), toFunc_(elementSelector_))
    
End Function


Private Function makeDictionary_(iterator_ As IEnumVARIANT, keySelector_ As IFunc, elmSelector_ As IFunc) As Dictionary
    
    Dim dict_ As New Dictionary
    
    
    Dim item_
    
    Do While xCom.CallNext_EnumVariant(iterator_, out_item_:=item_)
        
        
        Dim key_
        
        If keySelector_ Is Nothing _
            Then VariantCopy key_, item_ _
            Else VariantCopy key_, keySelector_.xExec01(item_)
        
        
        Dim element_
        
        If elmSelector_ Is Nothing _
            Then VariantCopy element_, item_ _
            Else VariantCopy element_, elmSelector_.xExec01(item_)
        
        
        
        '�v�f�������ɒǉ�����B
        '�L�[�������I�u�W�F�N�g�Ȃ�A�S�v�f�𕶎��񉻂��ăV���A���C�Y����B
        
        dict_.Add Grouping.ToAyKey(key_), element_
        
        
    Loop
    
    
    Set makeDictionary_ = dict_
    
End Function







' ---------------------------------------------------------------------------------








'���ʎ擾�n ---------------------------------------------------------------------------------



'�v�f�̌���Ԃ��B�\�[�X���J�E���g�\�Ȃ�A���ڂ̃J�E���g��Ԃ��B

Public Property Get Count() As Long
    
    If IsBlank Then Exit Property
    
    
    Select Case True
        
        Case TypeOf enumerableSource_ Is Enumerator
            
            Dim i
            For Each i In Me
                Count = Count + 1
            Next
            
            
        Case HasCountable
            
            Count = asObject_(enumerableSource_).Count
            
            
    End Select
    
End Property




'�v�f�̒l�����v����B�v�f�����l�ȊO�Ȃ疳������B

Public Function Sum() As Variant
    
    Dim i
    For Each i In Me
        
        If IsNumeric(i) Then Sum = Sum + i
        
    Next
    
End Function




' ---------------------------------------------------------------------------------






'�� ---------------------------------------------------------------------------------


Public Function ForEach(expression_) As Enumerator
    Set ForEach = Me
    
    Dim exp_ As IFunc
    Set exp_ = toFunc_(expression_)
    
    If exp_ Is Nothing Then Exit Function
    
    Dim i
    For Each i In Me
        
        exp_.xExec01 i
        
    Next
    
End Function


' ---------------------------------------------------------------------------------






'���̑� ---------------------------------------------------------------------------------




'�\�[�X���Z�b�g����Ă��Ȃ���ΐ^��Ԃ��B

Public Property Get IsBlank() As Boolean
    
    IsBlank = enumerableSource_ Is Nothing
    
End Property




'�\�[�X���J�E���g�\�i.Count�v���p�e�B�����j�Ȃ�ΐ^��Ԃ��B�������AEnumerator �͏����B

Public Function HasCountable() As Boolean
    
    If TypeOf enumerableSource_ Is Enumerator Then Exit Function
    
    On Error Resume Next
    
    HasCountable = IsNumeric(asObject_(enumerableSource_).Count)
    
End Function




' ---------------------------------------------------------------------------------











' �����p =============================================================================================




'�ЂȌ` EnumOperatorProcs.xxx �f���Q�[�g���C���X�^���X�������A�\�[�X�I�u�W�F�N�g�i EnumVariant ���񋓉\�ȃI�u�W�F�N�g�j���Z�b�g���ĕԂ��B

Private Function instantiateOperator_(operation_ As Delegate, srcEnum_ As IUnknown) As Delegate
    
    
    Dim newOp_ As EnumOperatorProcs
    
    Set newOp_ = operation_.Target.Instantiate
    
    
    Dim newDelegate_ As Delegate
    
    Set newDelegate_ = operation_.Clone
    
    
    Set newOp_.SourceIterator = srcEnum_
    
    Set instantiateOperator_ = newDelegate_.SetMethod(newOp_)
    
End Function





'�L���X�g�@unknown �� object

Private Function asObject_(unk_ As IUnknown) As Object
    
    Set asObject_ = unk_
    
End Function





'������R�[�h�Ȃ� Delegate �ɂ��ĕԂ��B�A�h���X�̉\��������� FuncPointer �ɂ��ĕԂ��B
'�@Delegate �Ȃǂ� IFunc �Ȃ炻�̂܂ܕԂ��Bmissing �Ȃ� Nothing ��Ԃ��B

Private Function toFunc_(expression_) As IFunc
    
    If IsMissing(expression_) Then Exit Function
    
    Select Case True
        
        Case TypeOf expression_ Is IFunc:                   Set toFunc_ = expression_
        
        Case varType(expression_) = vbString:               Set toFunc_ = Capt(CStr(expression_))
        
        Case varType(expression_) = varType(cNullPointer):  Set toFunc_ = FuncPointer.CNew(CLngPtr(expression_))
        
    End Select
    
End Function








