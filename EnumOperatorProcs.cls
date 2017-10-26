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



'ＬＩＮＱ的オペレータの処理を記述する。

'　また、Enumerator の operationDelegate_ デリゲートにオペレータメソッドとともにインスタンスがセットされ、
'　　オペレータ処理に必要な SourceIterator, Expression, Variable 変数を提供する。








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
Private Declare PtrSafe Function VariantCopyInd Lib "oleaut32" (dst_, src_) As Long

'--------------------









'メンバ宣言 ============================================================




Public SourceIterator   As IUnknown 'IEnumVARIANT or source object

Public Expression       As Variant  'オペレーションデリゲートなど、列挙単位で変化しないものを保持する。

Public Variable         As Variant  'カウンター変数など、列挙単位で変化する








'初期化／インスタンス化 ============================================================





'オブジェクトインスタンス化
'　SourceIterator 以外に対して、浅いコピーを行う（配列が格納されている場合は値コピーだが）。

Public Function Instantiate() As EnumOperatorProcs
    
    Set Instantiate = New EnumOperatorProcs
    
    Instantiate.SetParams Expression, Variable ' variable のほうも、配列領域のコピーが必要
    
End Function



' expression（式:IFunc）とvariable（変数）をセットする。

Public Function SetParams(expressionParam_, variableParam_) As EnumOperatorProcs
    
    Set SetParams = Me
    
    VariantCopyInd Expression, expressionParam_
    
    VariantCopyInd Variable, variableParam_
    
End Function



'式:IFunc をセットする。インデックスも指定した場合は、配列の要素に対して代入する。

Public Function SetExpression(expression_ As IFunc, Optional index_& = -1) As EnumOperatorProcs
    Set SetExpression = Me
    
    If index_ >= 0 _
        Then Set Expression(index_) = expression_ _
        Else Set Expression = expression_
    
End Function












'イテレーション =============================================================

'イテレーション処理は、列挙途中なら真を返す。列挙が終了していれば偽を返す。




' Ary の列挙子。EnumVariant の列挙処理を記述。

Public Function OpIteratorAry(operationDelegate_, ByRef out_item_) As Variant
    
    Dim srcAry_ As Ary: Set srcAry_ = SourceIterator
    
    Dim index_&:        index_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = srcAry_.Acquire(index_, out_item_:=out_item_)
    
    If Not isRetrieved_ Then Exit Function
    
    
    Variable = index_ + 1
    
    OpIteratorAry = isRetrieved_
    
End Function






' Enumerable.Range() の生成子。

Public Function OpGeneratorRange(operationDelegate_, ByRef out_item_) As Variant
    
    Dim i:              i = Variable(0)
    
    Dim span_:          span_ = Variable(1)
    
    Dim remaining_&:    remaining_ = Variable(2)
    
    
    If remaining_ <= 0 Then Exit Function
    
    out_item_ = i
    
    
    Variable(0) = i + span_
    
    Variable(2) = remaining_ - 1
    
    OpGeneratorRange = True
    
End Function







'オペレーション =============================================================

'イテレーション処理は、列挙途中なら真を返す。列挙が終了していれば偽を返す。






' qSelect( selector )

' selector
'　element : variant =>> result : variant

Public Function OpSelect(operationDelegate_, ByRef out_item_) As Variant
    
    Dim selector_   As IFunc:   Set selector_ = Expression
    
    
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
        
        
        VariantCopyInd out_item_, selector_.xExec01(out_item_)
        
        
        OpSelect = True
        
        Exit Function
        
    Loop
    
    
End Function





' qWhere( predicate )

' predicate
'　element : variant =>> isPassed : boolean

Public Function OpWhere(operationDelegate_, ByRef out_item_) As Variant
    
    Dim predicate_  As IFunc:   Set predicate_ = Expression
    
    
    Dim isRetrieved_ As Boolean
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
        
        isRetrieved_ = predicate_.xExec01(out_item_)
        
        If isRetrieved_ Then Exit Do
        
    Loop
    
    
    OpWhere = isRetrieved_
    
End Function





' qTake( count )

' count : long

Public Function OpTake(operationDelegate_, ByRef out_item_) As Variant
    
    Dim counter_&:  counter_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_) And counter_ > 0
    
    
    Variable = counter_ - 1
    
    OpTake = isRetrieved_
    
End Function





' qSkip( count )

' count : long

Public Function OpSkip(operationDelegate_, ByRef out_item_) As Variant
    
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





' qSpan( count )

' count : long

Public Function OpSpan(operationDelegate_, ByRef out_item_) As Variant
    
    Dim counter_&:  counter_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_)
        
        
        isRetrieved_ = counter_ <= 0
        
        If isRetrieved_ Then Exit Do
        
        
        counter_ = counter_ - 1
        
    Loop
    
    
    OpSpan = isRetrieved_
    
End Function








' qSelectMany( [collectionSelector][, resultSelector] )

' collection selector   （省略時は、流れてくる値そのものを enumerable とみなし展開する）
'　outer_element : variant =>> next_source : Enumerator     ※現状オブジェクトのみ、array は受け付けない

' result selector　     （省略時は、流れてくる値をそのまま次へ流す）
'　outer_element, inner_element =>> result : variant

Public Function OpSelectMany(operationDelegate_, ByRef out_innerItem_) As Variant
    
    Dim outerIterator_  As IEnumVARIANT:    Set outerIterator_ = SourceIterator
    Dim innerIterator_  As IEnumVARIANT:    Set innerIterator_ = Variable(0)
    Dim outerItem_:                         Set outerItem_ = Variable(1)
    
    Dim colSelector_    As IFunc:           Set colSelector_ = Expression(0)
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression(1)
    
    
    '列挙再開
    
    If Not innerIterator_ Is Nothing Then GoTo OnInner
    
    
    
    '外側の列挙
    
    Do While xCom.CallNext_EnumVariant(outerIterator_, out_item_:=outerItem_)
        
        If Not colSelector_ Is Nothing Then
            
            VariantCopyInd outerItem_, colSelector_.xExec01(outerItem_)
            
        End If
        
        
        '内側の EnumVariant を取得する。
        
        Set innerIterator_ = xCom.GetEnumVariant(outerItem_)
        
'        If innerIterator_ Is Nothing Then err.Raise xxx '列挙可能ではなかった
        
        
OnInner: '内側の列挙
        
        Do While xCom.CallNext_EnumVariant(innerIterator_, out_item_:=out_innerItem_)
            
            If Not resSelector_ Is Nothing Then
                
                VariantCopyInd out_innerItem_, resSelector_.xExec02(outerItem_, out_innerItem_)
                
            End If
            
            
            Set Variable(0) = innerIterator_
            
            Set Variable(1) = outerItem_
            
            OpSelectMany = True
            
            Exit Function
            
        Loop
        
    Loop
    
    
End Function









' qGroupBy( keySelector[, elementSelector][, resultSelector] )

' keySelector       （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> key : variant

' elementSelector   （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> element_result : variant

' resultSelector　GroupBy().Select() と同じこと。必要だろうか？
'　key : variant, result : enumerable =>> result : variant
'※省略すると Grouping が列挙され、省略しなければ Grouping の全要素が key とともに列挙される。


'ステップ１：辞書構築→第一要素を流す

Public Function OpGroupBy(operationDelegate_, ByRef out_item_) As Variant
    
    Dim keySelector_    As IFunc:   Set keySelector_ = Expression(0)
    Dim elmSelector_    As IFunc:   Set elmSelector_ = Expression(1)
    Dim resSelector_    As IFunc:   Set resSelector_ = Expression(2)
    
    
    
    'グルーピング結果を構築する。
    
    Dim groups_ As Dictionary
    
    Set groups_ = Enumerable(SourceIterator).ToLookup(keySelector_, elmSelector_)
    
    Set Variable(1) = groups_
    
    
    
    'グルーピングの EnumVariant を取得し、今までの enumerableSource を指すものと差し替える。
    
    Set SourceIterator = xCom.GetEnumVariant(groups_)
    
    
    
    '次からの挙動をグルーピング列挙を行うオペレーターに差し替える。
    
    If resSelector_ Is Nothing _
        Then shiftOperation_ operationDelegate_, "OpIterateGroup" _
        Else shiftOperation_ operationDelegate_, "OpIterateGroupWithSelector"
    
    
    
    '列挙をひとつ実行する。
    
    OpGroupBy = asIFunc_(operationDelegate_).xExec02(operationDelegate_, out_item_)
    
    
End Function



'ステップ２：グループの列挙（グループを返す）　※ resultSelector_ が渡されていない場合

Public Function OpIterateGroup(operationDelegate_, ByRef out_item_) As Variant
    
    Dim groupIterator_  As IEnumVARIANT:    Set groupIterator_ = SourceIterator
    Dim groups_         As Dictionary:      Set groups_ = Variable(1)
    
    
    Dim key_
    
    Do While xCom.CallNext_EnumVariant(groupIterator_, out_item_:=key_)
        
        Set out_item_ = groups_(key_)
        
        OpIterateGroup = True
        
        Exit Function
        
    Loop
    
End Function



'ステップ２´：グループ内要素の列挙（グループ内の要素を返す）　※ resultSelector_ が渡されていた場合

Public Function OpIterateGroupWithSelector(operationDelegate_, ByRef out_innerItem_) As Variant
    
    Dim groupIterator_  As IEnumVARIANT:    Set groupIterator_ = SourceIterator
    Dim innerIterator_  As IEnumVARIANT:    Set innerIterator_ = Variable(0)
    
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression(2)
    
    Dim groups_         As Dictionary:      Set groups_ = Variable(1)
    Dim originalKey_:                       VariantCopy originalKey_, Variable(2)
    Dim dictinalyKey_
    
    
    '列挙再開
    
    If Not innerIterator_ Is Nothing Then GoTo OnInner
    
    
    
    '外側の列挙
    
    Dim outerItem_
    
    Do While xCom.CallNext_EnumVariant(groupIterator_, out_item_:=dictinalyKey_)
        
        
        '内側の EnumVariant を取得する。
        
        Dim group_ As Grouping:     Set group_ = groups_(dictinalyKey_)
        
        Set innerIterator_ = xCom.GetEnumVariant(group_.Elements)
        
        VariantCopyInd originalKey_, group_.Key
        
        
OnInner: '内側の列挙
        
        Do While xCom.CallNext_EnumVariant(innerIterator_, out_item_:=out_innerItem_)
            
            
            'セレクターを通して値を返す。
            
            VariantCopy out_innerItem_, resSelector_.xExec02(originalKey_, out_innerItem_)
            
            
            VariantCopyInd Variable(2), originalKey_
            
            Set Variable(0) = innerIterator_
            
            OpIterateGroupWithSelector = True
            
            Exit Function
            
        Loop
        
    Loop
    
    
End Function










' qScan( func_[, resultSelector] )

' func_
'　accumulate : variant, element : variant =>> accumulate : variant

' resultSelector　Scan().Select() と同じこと。必要だろうか？
'　accumulate : variant =>> result : variant


Public Function OpScanFirstStep(operationDelegate_, ByRef out_item_) As Variant
    
    Dim accumulate_
    
    
    ' seed を伴わない場合の最初の値を取得する。
    
    If Not xCom.CallNext_EnumVariant(SourceIterator, out_item_:=accumulate_) Then Exit Function
    
    
    VariantCopyInd Variable, accumulate_
    
    
    OpScanFirstStep = OpScanFirstStep2(operationDelegate_, out_item_:=out_item_)
    
End Function


Public Function OpScanFirstStep2(operationDelegate_, ByRef out_item_) As Variant
    
    Dim func_           As IFunc:   Set func_ = Expression(0)
    Dim resSelector_    As IFunc:   Set resSelector_ = Expression(1)
    
    Dim accumulate_:                VariantCopy accumulate_, Variable
    
    
    '最初の値を result selector に通して返す。
    
    If Not resSelector_ Is Nothing _
        Then VariantCopy out_item_, resSelector_.xExec01(accumulate_) _
        Else VariantCopy out_item_, accumulate_
    
    
    shiftOperation_ operationDelegate_, "OpScan"
    
    OpScanFirstStep2 = True
    
End Function


Public Function OpScan(operationDelegate_, ByRef out_item_) As Variant
    
    Dim func_           As IFunc:   Set func_ = Expression(0)
    Dim resSelector_    As IFunc:   Set resSelector_ = Expression(1)
    
    Dim accumulate_:                VariantCopy accumulate_, Variable
    
    Dim element_
    
    
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=element_)
        
        
        '前の値と今回の値を func にかける。
        
        VariantCopyInd Variable, func_.xExec02(accumulate_, element_)
        
        
        '結果を result selecor に通して返す。
        
        If Not resSelector_ Is Nothing _
            Then VariantCopy out_item_, resSelector_.xExec01(Variable) _
            Else VariantCopy out_item_, accumulate_
        
        
        OpScan = True
        
        Exit Function
    Loop
    
    
End Function


    









' qConcat/qUnion/qDistinct



'重複なしの和集合を行う前の辞書準備

Public Function OpUnion(operationDelegate_, ByRef out_item_) As Variant
    
    Set Variable(0) = SourceIterator '第１イテレータ
    
    Set Variable(2) = New Dictionary '重複除去のための辞書をセットする。
    
    shiftOperation_ operationDelegate_, "OpUnionStepFirst"
    
    OpUnion = OpUnionStepFirst(operationDelegate_, out_item_:=out_item_)
    
End Function

Public Function OpUnionAll(operationDelegate_, ByRef out_item_) As Variant
    
    Set Variable(0) = SourceIterator '第１イテレータ
    
    shiftOperation_ operationDelegate_, "OpUnionStepFirst"
    
    OpUnionAll = OpUnionStepFirst(operationDelegate_, out_item_:=out_item_)
    
End Function


'１つめのソースが終わるまでこれを経由する。

Public Function OpUnionStepFirst(operationDelegate_, ByRef out_item_) As Variant
    
    
    '列挙
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = OpUnionBody(operationDelegate_, out_item_:=out_item_)
    
    
    
    '列挙が完了した場合は、次のソースがあれば移る。
    
    If Not isRetrieved_ Then
        
        Dim nextSource_ As Object:  Set nextSource_ = Variable(1)
        
        If nextSource_ Is Nothing Then Exit Function
        
        
        '第２イテレータ
        
        Set Variable(0) = Enumerable(nextSource_).NewEnum ' next iterator_
        
        shiftOperation_ operationDelegate_, "OpUnionBody"
        
        
        isRetrieved_ = OpUnionBody(operationDelegate_, out_item_:=out_item_)
        
    End If
    
    
    OpUnionStepFirst = isRetrieved_
    
End Function


'イテレーション本体

Public Function OpUnionBody(operationDelegate_, ByRef out_item_) As Variant
    
    Dim iterator_   As IEnumVARIANT:    Set iterator_ = Variable(0)
    
    Dim nextSource_ As Object:          Set nextSource_ = Variable(1)
    
    Dim distincts_  As Dictionary:      Set distincts_ = Variable(2)
    
    
    
    Dim isRetrieved_ As Boolean
    
    Do While xCom.CallNext_EnumVariant(iterator_, out_item_:=out_item_)
        
        
        isRetrieved_ = distincts_ Is Nothing
        
        If isRetrieved_ Then Exit Do
        
        
        isRetrieved_ = Not isKeyExisted_(distincts_, out_item_)
        
        If isRetrieved_ Then Exit Do
        
        
    Loop
    
    
    OpUnionBody = isRetrieved_
    
End Function



Private Function isKeyExisted_(distincts_ As Dictionary, element_) As Boolean
    
    Dim key_:   key_ = Grouping.ToAyKey(element_)
    
    isKeyExisted_ = distincts_.Exists(key_)
    
    If isKeyExisted_ Then Exit Function
    
    
    distincts_.Add key_, Empty
    
End Function










' qZip

' resultSelector
'　first_source : variant, second_source : variant =>> result : variant

Public Function OpZip(operationDelegate_, ByRef out_item_) As Variant
    
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression
    
    Dim firstIterator_  As IEnumVARIANT:    Set firstIterator_ = SourceIterator
    
    Dim secondIterator_ As IEnumVARIANT:    Set secondIterator_ = Variable(0)
    
    Dim secondSource_   As Object:          Set secondSource_ = Variable(1)
    
    
    
    If secondIterator_ Is Nothing Then
        
        Set secondIterator_ = Enumerable(secondSource_).NewEnum
        
        Set Variable(0) = secondIterator_
        
    End If
    
    
    Do
        
        Dim firstElement_
        
        If Not xCom.CallNext_EnumVariant(firstIterator_, out_item_:=firstElement_) Then Exit Function
        
        Dim secondElement_
        
        If Not xCom.CallNext_EnumVariant(secondIterator_, out_item_:=secondElement_) Then Exit Function
        
        
        If Not resSelector_ Is Nothing _
            Then VariantCopy out_item_, resSelector_.xExec02(firstElement_, secondElement_) _
            Else VariantCopy out_item_, resultSelectorDefaultOfZip_(firstElement_, secondElement_)
        
        
        OpZip = True
        
        Exit Function
        
    Loop
    
    
End Function


' result selector 省略時のデフォルト

Private Function resultSelectorDefaultOfZip_(firstElement_, secondElement_) As Variant
    
    Dim c_ As New Collection
    
    c_.Add Grouping.AnonymousIdentify, "ay"
    
    c_.Add firstElement_, "first"
    
    c_.Add secondElement_, "second"
    
    Set resultSelectorDefaultOfZip_ = c_
    
End Function







' qLetIf

' comparer
'　thisElement : variant, otherElement : variant =>> isEstablished : boolean

' selector
'　source : variant =>> conversion : variant

Public Function OpLetIf(operationDelegate_, ByRef out_item_) As Variant
    
    Dim selector_   As IFunc:   Set selector_ = Expression(1)
    
    
    If xCom.CallNext_EnumVariant(SourceIterator, out_item_:=out_item_) Then
        
        If Not selector_ Is Nothing Then
            
            VariantCopyInd out_item_, selector_.xExec01(out_item_)
            
        End If
        
        VariantCopyInd Variable, out_item_
        
        
        shiftOperation_ operationDelegate_, "OpLetIfNext"
        
        OpLetIf = True
        
    End If
    
End Function


Public Function OpLetIfNext(operationDelegate_, ByRef out_item_) As Variant
    
    Dim comparer_   As IFunc:   Set comparer_ = Expression(0)
    Dim selector_   As IFunc:   Set selector_ = Expression(1)
    
    Dim prevElement_:           VariantCopy prevElement_, Variable
    Dim thisElement_
    VariantCopy out_item_, Variable
    
    Do While xCom.CallNext_EnumVariant(SourceIterator, out_item_:=thisElement_)
        
        If comparer_.xExec02(prevElement_, thisElement_) Then
            
            If Not selector_ Is Nothing Then
                
                VariantCopyInd thisElement_, selector_.xExec01(thisElement_)
                
            End If
            
            VariantCopyInd Variable, thisElement_
            
            VariantCopyInd out_item_, thisElement_
            
            
            OpLetIfNext = True
            
            Exit Function
            
        End If
        
    Loop
    
End Function









Public Function IsMax(other_, this_) As Variant
    
    IsMax = this_ > other_
    
End Function

Public Function IsMin(other_, this_) As Variant
    
    IsMin = this_ < other_
    
End Function













'要素を辞書に追加する。

Public Function RegistDictionary(dict_, key_, element_) As Variant
    
    Dim elements_ As Dictionary:    Set elements_ = dict_
    
    elements_.Add Grouping.ToAyKey(key_), element_
    
End Function






'要素をグループコレクションに追加する。

Public Function RegistGroupDictionary(dict_, key_, element_) As Variant
    
    Dim groups_ As Dictionary:  Set groups_ = dict_
    
    getGrouping_(groups_, key_).Elements.Add element_
    
End Function


'要素をグループコレクションに追加する｡
'　辞書から Grouping を取り出す。登録されていなければ新規 Grouping を登録する。

Private Function getGrouping_(groups_ As Dictionary, key_) As Grouping
    
    
    Dim dictKey_:   dictKey_ = Grouping.ToAyKey(key_)
    
    
    ' key_ が辞書に登録されていればそれを返す。
    
    On Error Resume Next
    
    Set getGrouping_ = groups_(dictKey_)
    
    If Not getGrouping_ Is Nothing Then Exit Function
    
    
    
    '登録されていなければ新規作成、登録、そして返す。
    
    Set getGrouping_ = New Grouping
    
    If IsObject(key_) _
        Then Set getGrouping_.Key = key_ _
        Else getGrouping_.Key = key_
    
    Set groups_(dictKey_) = getGrouping_
    
    
End Function














'次からのオペレーターを差し替える。

Private Sub shiftOperation_(operationDelegate_, methodName_$)
    
    Dim eop_ As Delegate:    Set eop_ = operationDelegate_
    
    eop_.SetMethod methodName_:=methodName_
    
End Sub


' IFunc にキャストする。

Private Function asIFunc_(operationDelegate_) As IFunc
    
    Set asIFunc_ = operationDelegate_
    
End Function





