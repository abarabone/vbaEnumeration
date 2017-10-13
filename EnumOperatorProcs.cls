VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumOperatorProcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    ByVal pvInstance_ As LongPtr, _
    ByVal oVft_ As LongPtr, ByVal cc_ As Long, _
    ByVal vtReturn_ As Integer, _
    ByVal cActuals_ As Long, ByRef valueTypeTop_ As Integer, ByRef argPtrTop_ As LongPtr, _
    ByRef pvargResult_ As Variant _
) As Long

Private Const cStdCall& = 4


Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (dst_, src_) As Long

'--------------------









'�����o�錾 ============================================================




Public SourceIterator   As IUnknown 'IEnumVARIANT or source object

Public Expression       As Variant  '�I�y���[�V�����f���Q�[�g�ȂǁA�񋓒P�ʂŕω����Ȃ����̂�ێ�����B

Public Variable         As Variant  '�J�E���^�[�ϐ��ȂǁA�񋓒P�ʂŕω�����








'�������^�C���X�^���X�� ============================================================





'�I�u�W�F�N�g�C���X�^���X��
'�@SourceIterator �ȊO�ɑ΂��āA�󂢃R�s�[���s���i�z�񂪊i�[����Ă���ꍇ�͒l�R�s�[�����j�B

Public Function Instantiate() As EnumOperatorProcs
    
    Set Instantiate = New EnumOperatorProcs
    
    Instantiate.SetParams Expression, Variable ' variable �̂ق����A�z��̈�̃R�s�[���K�v
    
End Function



' expression�i��:IFunc�j��variable�i�ϐ��j���Z�b�g����B

Public Function SetParams(expressionParam_, variableParam_) As EnumOperatorProcs
    
    Set SetParams = Me
    
    VariantCopy Expression, expressionParam_
    
    VariantCopy Variable, variableParam_
    
End Function



'��:IFunc ���Z�b�g����B�C���f�b�N�X���w�肵���ꍇ�́A�z��̗v�f�ɑ΂��đ������B

Public Function SetExpression(expression_ As IFunc, Optional index_& = -1) As EnumOperatorProcs
    Set SetExpression = Me
    
    If index_ >= 0 _
        Then Set Expression(index_) = expression_ _
        Else Set Expression = expression_
    
End Function












'�C�e���[�V���� =============================================================

'�C�e���[�V���������́A�񋓓r���Ȃ�^��Ԃ��B�񋓂��I�����Ă���΋U��Ԃ��B




' Ary �̗񋓎q�BEnumVariant �̗񋓏������L�q�B

Public Function OpIteratorAry(operator_, ByRef out_item_) As Variant
    
    Dim srcAry_ As Ary: Set srcAry_ = SourceIterator
    
    Dim index_&:        index_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = srcAry_.Acquire(index_, out_item_:=out_item_)
    
    If Not isRetrieved_ Then Exit Function
    
    
    Variable = index_ + 1
    
    OpIteratorAry = isRetrieved_
    
End Function









'�I�y���[�V���� =============================================================

'�C�e���[�V���������́A�񋓓r���Ȃ�^��Ԃ��B�񋓂��I�����Ă���΋U��Ԃ��B






' qSelect( selector )

' selector
'�@element : variant =>> result : variant

Public Function OpSelect(operator_, ByRef out_item_) As Variant
    
    Dim selector_   As IFunc:   Set selector_ = Expression
    
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
    
    If isRetrieved_ Then
        
        VariantCopy out_item_, selector_.xExec01(out_item_)
        
    End If
    
    
    OpSelect = isRetrieved_
    
End Function





' qWhere( predicate )

' predicate
'�@element : variant =>> isSend : boolean

Public Function OpWhere(operator_, ByRef out_item_) As Variant
    
    Dim predicate_  As IFunc:   Set predicate_ = Expression
    
    
    Dim isRetrieved_ As Boolean
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
        
        isRetrieved_ = predicate_.xExec01(out_item_)
        
        If isRetrieved_ Then Exit Do
        
    Loop
    
    
    OpWhere = isRetrieved_
    
End Function





' xTake( count )

' count : long

Public Function OpTake(operator_, ByRef out_item_) As Variant
    
    Dim counter_&:  counter_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_) And counter_ > 0
    
    
    Variable = counter_ - 1
    
    OpTake = isRetrieved_
    
End Function





' xSkip( count )

' count : long

Public Function OpSkip(operator_, ByRef out_item_) As Variant
    
    Dim counter_&:  counter_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
        
        
        isRetrieved_ = counter_ <= 0
        
        If isRetrieved_ Then Exit Do
        
        
        counter_ = counter_ - 1
        
    Loop
    
    
    Variable = counter_
    
    OpSkip = isRetrieved_
    
End Function





' xSpan( count )

' count : long

Public Function OpSpan(operator_, ByRef out_item_) As Variant
    
    Dim counter_&:  counter_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
        
        
        isRetrieved_ = counter_ <= 0
        
        If isRetrieved_ Then Exit Do
        
        
        counter_ = counter_ - 1
        
    Loop
    
    
    OpSpan = isRetrieved_
    
End Function






' xSelectMany( [collectionSelector][, resultSelector] )

' collection selector   �i�ȗ����́A����Ă���l���̂��̂� enumerable �Ƃ݂Ȃ��W�J����j
'�@outer_element : variant =>> next_source : Enumerator

' result selector�@     �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@outer_element, inner_element =>> result : variant

Public Function OpSelectMany(operator_, ByRef out_innerItem_) As Variant
    
    Dim outerIterator_  As IEnumVARIANT:    Set outerIterator_ = SourceIterator
    Dim innerIterator_  As IEnumVARIANT:    Set innerIterator_ = Variable
    
    Dim colSelector_    As IFunc:           Set colSelector_ = Expression(0)
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression(1)
    
    
    Dim isRetrieved_ As Boolean
    
    Do
        
        If innerIterator_ Is Nothing Then
            
            '�O���̗�
            
            Dim outerItem_
            
            isRetrieved_ = iterate_(outerIterator_, colSelector_, out_item_:=outerItem_)
            
            If Not isRetrieved_ Then Exit Function '�񋓏I��
            
            
            '������ EnumVariant ���擾����B
            
            Set innerIterator_ = xCom.GetEnumVariant(outerItem_)
            
            Set Variable = innerIterator_
            
            If innerIterator_ Is Nothing Then Exit Function '�񋓉\�ł͂Ȃ�����
            
        End If
        
        
        '�����̗�
        
        isRetrieved_ = iterate_(innerIterator_, resSelector_, out_item_:=out_innerItem_)
        
        Set innerIterator_ = Nothing
        
    Loop Until isRetrieved_
    
    
    OpSelectMany = isRetrieved_
    
End Function

Private Function iterate_(iterator_ As IEnumVARIANT, selector_ As IFunc, ByRef out_item_) As Boolean
    
    
    iterate_ = xCom.CallNext_EnumVariant(iterator_, out_item_:=out_item_)
    
    If Not iterate_ Then Exit Function
    
    
    If selector_ Is Nothing Then Exit Function
    
    VariantCopy out_item_, selector_.xExec01(out_item_)
    
    
End Function







' qGroupBy( keySelector[, elementSelector][, resultSelector] )

' keySelector       �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@element : variant =>> key : variant

' elementSelector   �i�ȗ����́A����Ă���l�����̂܂܎��֗����j
'�@element : variant =>> element_result : variant

' resultSelector�@GroupBy().Select() �Ɠ������ƁB�K�v���낤���H
'�@key : variant, result : enumerable =>> result : variant
'���ȗ������ Grouping ���񋓂���A�ȗ����Ȃ���� Grouping �̑S�v�f�� key �ƂƂ��ɗ񋓂����B


'�X�e�b�v�P�F�����\�z�����v�f�𗬂�

Public Function OpGroupBy(operator_, ByRef out_item_) As Variant
    
    Dim keySelector_    As IFunc:   Set keySelector_ = Expression(0)
    Dim elmSelector_    As IFunc:   Set elmSelector_ = Expression(1)
    Dim resSelector_    As IFunc:   Set resSelector_ = Expression(2)
    
    
    '�O���[�s���O���ʂ��\�z����B
    
    Dim groups_ As Dictionary
    
    Set groups_ = Enumerable(SourceIterator).ToLookUp(keySelector_, elmSelector_)
    
    Set Variable(0) = groups_
    
    
    '�O���[�s���O�� EnumVariant ���擾���A���܂ł� enumerableSource ���w�����̂ƍ����ւ���B
    
    Set SourceIterator = xCom.GetEnumVariant(groups_)
    
    
    '������̋������O���[�s���O�񋓂��s���I�y���[�^�[�ɍ����ւ���B
    
    Dim op_ As Delegate:    Set op_ = operator_
    
    op_.SetMethod methodName_:=IIf(resSelector_ Is Nothing, "OpIterateGroup", "OpIterateGroupWithSelector")
    
    
    '�񋓂��ЂƂ��s����B
    
    Dim nextOperation_ As IFunc:    Set nextOperation_ = operator_
    
    OpGroupBy = nextOperation_(operator_, out_item_)
    
End Function


'�X�e�b�v�Q�F�O���[�v�̗񋓁i�O���[�v��Ԃ��j

Public Function OpIterateGroup(operator_, ByRef out_item_) As Variant
    
    Dim groupIterator_  As IEnumVARIANT:    Set groupIterator_ = SourceIterator
    Dim groups_         As Dictionary:      Set groups_ = Variable(0)
    
    
    Dim isRetrieved_ As Boolean
    
    
    Dim key_
    
    isRetrieved_ = xCom.CallNext_EnumVariant(groupIterator_, out_item_:=key_)
    
    If Not isRetrieved_ Then Exit Function
    
    
    Set out_item_ = groups_(key_)
    
    OpIterateGroup = isRetrieved_
    
End Function


'�X�e�b�v�Q�F�O���[�v���v�f�̗񋓁i�O���[�v���̗v�f��Ԃ��j

Public Function OpIterateGroupWithSelector(operator_, ByRef out_innerItem_) As Variant
    
    Dim groupIterator_  As IEnumVARIANT:    Set groupIterator_ = SourceIterator
    Dim innerIterator_  As IEnumVARIANT:    Set innerIterator_ = Variable(2)
    
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression(2)
    
    Dim groups_         As Dictionary:      Set groups_ = Variable(0)
    Dim key_
    
    
    Dim isRetrieved_ As Boolean
    
    Do
        
        If innerIterator_ Is Nothing Then
            
            '�O���̗�
            
            isRetrieved_ = xCom.CallNext_EnumVariant(groupIterator_, out_item_:=key_)
            
            If Not isRetrieved_ Then Exit Function '�񋓏I��
            
            
            '������ EnumVariant ���擾����B
            
            Dim group_ As Grouping:     Set group_ = groups_(key_)
            
            Set innerIterator_ = xCom.GetEnumVariant(group_.Elements)
            
            Set Variable(2) = innerIterator_
            
            VariantCopy Variable(1), group_.Key
            
        End If
        
        
        '�����̗�
        
        isRetrieved_ = xCom.CallNext_EnumVariant(innerIterator_, out_item_:=out_innerItem_)
        
        Set innerIterator_ = Nothing
        
    Loop Until isRetrieved_
    
    
    '�Z���N�^�[��ʂ��Ēl��Ԃ��B
    
    VariantCopy key_, Variable(1)
    
    VariantCopy out_innerItem_, resSelector_.xExec02(key_, out_innerItem_)
    
    OpIterateGroupWithSelector = isRetrieved_
    
End Function














