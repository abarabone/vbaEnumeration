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

'　また、Enumerator の operation_ デリゲートにオペレータメソッドとともにインスタンスがセットされ、
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

Public Function OpIteratorAry(operator_, ByRef out_item_) As Variant
    
    Dim srcAry_ As Ary: Set srcAry_ = SourceIterator
    
    Dim index_&:        index_ = Variable
    
    
    Dim isRetrieved_ As Boolean
    
    isRetrieved_ = srcAry_.Acquire(index_, out_item_:=out_item_)
    
'    out_item_ = srcAry_(index_)
'
'    isRetrieved_ = (index_ < srcAry_.Count)
    
    If Not isRetrieved_ Then Exit Function
    
    
    Variable = index_ + 1
    
    OpIteratorAry = isRetrieved_
    
End Function









'オペレーション =============================================================

'イテレーション処理は、列挙途中なら真を返す。列挙が終了していれば偽を返す。






' qSelect( selector )

' selector
'　element : variant =>> result : variant

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
'　element : variant =>> isSend : boolean

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

' collection selector   （省略時は、流れてくる値そのものを enumerable とみなし展開する）
'　outer_element : variant =>> next_source : Enumerator

' result selector　     （省略時は、流れてくる値をそのまま次へ流す）
'　outer_element, inner_element =>> result : variant

Public Function OpSelectMany(operator_, ByRef out_innerItem_) As Variant
    
    Dim outerIterator_  As IEnumVARIANT:    Set outerIterator_ = SourceIterator
    Dim innerIterator_  As IEnumVARIANT:    Set innerIterator_ = Variable
    
    Dim colSelector_    As IFunc:           Set colSelector_ = Expression(0)
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression(1)
    
    
    Dim isRetrieved_ As Boolean
    
    Do
        
        If innerIterator_ Is Nothing Then
            
            '外側の列挙
            
            Dim outerItem_
            
            isRetrieved_ = iterate_(outerIterator_, colSelector_, out_item_:=outerItem_)
            
            If Not isRetrieved_ Then Exit Function '列挙終了
            
            
            '内側の EnumVariant を取得する。
            
            Set innerIterator_ = xCom.GetEnumVariant(outerItem_)
            
            Set Variable = innerIterator_
            
            If innerIterator_ Is Nothing Then Exit Function '列挙可能ではなかった
            
        End If
        
        
        '内側の列挙
        
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

' keySelector       （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> key : variant

' elementSelector   （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> element_result : variant

' resultSelector　GroupBy().Select() と同じこと。必要だろうか？
'　key : variant, result : enumerable =>> result : variant
'※省略すると Grouping が列挙され、省略しなければ Grouping の全要素が key とともに列挙される。


'ステップ１：辞書構築→第一要素を流す

Public Function OpGroupBy(operator_, ByRef out_item_) As Variant
    
    Dim keySelector_    As IFunc:   Set keySelector_ = Expression(0)
    Dim elmSelector_    As IFunc:   Set elmSelector_ = Expression(1)
    Dim resSelector_    As IFunc:   Set resSelector_ = Expression(2)
    
    
    'グルーピング結果を構築する。
    
    Dim groups_ As Dictionary
    
    Set groups_ = Enumerable(SourceIterator).ToLookup(keySelector_, elmSelector_)
    
    Set Variable(0) = groups_
    
    
    'グルーピングの EnumVariant を取得し、今までの enumerableSource を指すものと差し替える。
    
    Set SourceIterator = xCom.GetEnumVariant(groups_)
    
    
    '次からの挙動をグルーピング列挙を行うオペレーターに差し替える。
    
    Dim op_ As Delegate:    Set op_ = operator_
    
    op_.SetMethod methodName_:=IIf(resSelector_ Is Nothing, "OpIterateGroup", "OpIterateGroupWithSelector")
    
    
    '列挙をひとつ実行する。
    
    Dim nextOperation_ As IFunc:    Set nextOperation_ = operator_
    
    OpGroupBy = nextOperation_(operator_, out_item_)
    
End Function


'ステップ２：グループの列挙（グループを返す）

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


'ステップ２：グループ内要素の列挙（グループ内の要素を返す）

Public Function OpIterateGroupWithSelector(operator_, ByRef out_innerItem_) As Variant
    
    Dim groupIterator_  As IEnumVARIANT:    Set groupIterator_ = SourceIterator
    Dim innerIterator_  As IEnumVARIANT:    Set innerIterator_ = Variable(2)
    
    Dim resSelector_    As IFunc:           Set resSelector_ = Expression(2)
    
    Dim groups_         As Dictionary:      Set groups_ = Variable(0)
    Dim key_
    
    
    Dim isRetrieved_ As Boolean
    
    Do
        
        If innerIterator_ Is Nothing Then
            
            '外側の列挙
            
            isRetrieved_ = xCom.CallNext_EnumVariant(groupIterator_, out_item_:=key_)
            
            If Not isRetrieved_ Then Exit Function '列挙終了
            
            
            '内側の EnumVariant を取得する。
            
            Dim group_ As Grouping:     Set group_ = groups_(key_)
            
            Set innerIterator_ = xCom.GetEnumVariant(group_.Elements)
            
            Set Variable(2) = innerIterator_
            
            VariantCopyInd Variable(1), group_.Key
            
        End If
        
        
        '内側の列挙
        
        isRetrieved_ = xCom.CallNext_EnumVariant(innerIterator_, out_item_:=out_innerItem_)
        
        Set innerIterator_ = Nothing
        
    Loop Until isRetrieved_
    
    
    'セレクターを通して値を返す。
    
    VariantCopy key_, Variable(1)
    
    VariantCopy out_innerItem_, resSelector_.xExec02(key_, out_innerItem_)
    
    OpIterateGroupWithSelector = isRetrieved_
    
End Function















