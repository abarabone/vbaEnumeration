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





'列挙可能オブジェクト

'　基本的には、ＬＩＮＱ的なオペレータ操作を行うためのオブジェクトである。

'　ソースは、列挙可能オブジェクトチェインかそれ以外の列挙可能な集合（配列やコレクションなど）でなければならない。

'　列挙可能な集合をソースとして所持する場合、それをイテレーションする必要が生じる。
'　　そのために必ず列挙可能な集合を列挙子（イテレーション機能関数を持つ列挙可能オブジェクト）でくるむ。






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
    ByVal cActuals_ As Long, valueTypeTeop_ As Integer, argPtrTeop_ As LongPtr, _
    pvargResult_ As Variant _
) As Long

Const cStdCall& = 4


Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (ByRef dst_ As Any, ByRef src_ As Any) As Long
Private Declare PtrSafe Function VariantCopyInd Lib "oleaut32" (ByRef dst_ As Any, ByRef src_ As Any) As Long

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst_ As Any, ByRef src_ As Any, ByVal size_&)

Private Declare PtrSafe Function CoTaskMemAlloc Lib "Ole32" (ByVal byte_&) As LongPtr

Private Declare PtrSafe Sub CoTaskMemFree Lib "Ole32" (ByVal pMem_ As LongPtr)

'--------------------







'構造体定義 ============================================================




'自前 EnumVariant 構造体

Private Type EnumVariantStruct
    
    PVtable     As LongPtr
    Operator    As IFunc
    
    RefCount    As Long
    
End Type









'メンバ宣言 ============================================================




Private enumerableSource_   As IUnknown     '参照元となる列挙可能な集合。ソース。IEnumVariant を返せるもの（Enumerator 含む）

Private operation_          As Delegate     ' select や where といったオペレータ処理のデリゲートを格納する。

' operation_ デリゲートは処理時にはそのまま使用しない、ひな形である。
'　For Each 開始時にはその都度 EnumVariant が生成されるが、それに合わせてデリゲートもインスタンスを作成しなければならない。
'　なぜなら、デリゲートが所持する EnumOperatorProces オブジェクトにも状態の変化する変数 Variable があるからである。
'　さらに言うなら、そのオブジェクトは Expression デリゲートを所持する場合があるが、そちらは状態を持たない。







'関数定義 ============================================================





' For Each 等に EnumVariant を返す ------------------------------------------


' enumerableSource_ が未定義の場合は、Nothing が返る。

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    
    If IsBlank Then Exit Function
    
    
    Select Case True
        
        Case enumerableSource_ Is Nothing, TypeOf enumerableSource_ Is Ary
            
            Set NewEnum = xCom.CreateEnumVariant(instantiateOperator_(operation_, enumerableSource_))
            
            
        Case TypeOf enumerableSource_ Is Enumerator
            
            Set NewEnum = xCom.CreateEnumVariant(instantiateOperator_(operation_, asObject_(enumerableSource_).NewEnum))
            
            
        Case TypeOf enumerableSource_ Is Object
            
            Dim ienv_
            
            Const cIID_IEnumVARIANT& = -4
            
            xCom.Call_Invoke enumerableSource_, cIID_IEnumVARIANT, 0, 0, ienv_
            
            Set NewEnum = ienv_
            
            
        Case TypeOf enumerableSource_ Is IEnumVARIANT
            
            'ソースが EnumVariant の場合は、クローンでもしようと思うが、とりあえずはそのままかえす。
            
            Set NewEnum = enumerableSource_
            
            
    End Select
    
End Function


' ---------------------------------------------------------------------------









'列挙ソースとオペレーションデリゲートをセットする ------------------------------------------



'ソースオブジェクトが EnumVariant を返せない場合、デリゲートでイテレーションを記述し、ソースオブジェクトとともにセットする。

Public Function SetIterator(enumerableSourceObject_ As Object, operationFunction_ As Delegate) As Enumerator
    Set SetIterator = Me
    
    Set enumerableSource_ = enumerableSourceObject_
    
    Set operation_ = operationFunction_
    
End Function



'ソースとなる Enumrator とＬＩＮＱ的オペレータデリゲートをセットする。

Public Function SetOperator(sourceEnumerator_ As Enumerator, operationFunction_ As Delegate) As Enumerator
    Set SetOperator = Me
    
    Set enumerableSource_ = sourceEnumerator_
    
    Set operation_ = operationFunction_
    
End Function



'ソースとなるオブジェクトをセットする。EnumVariant も渡せる。

Public Function SetSource(enumerableSourceObject_ As IUnknown) As Enumerator
    Set SetSource = Me
    
    Set enumerableSource_ = enumerableSourceObject_
    
End Function



' ---------------------------------------------------------------------------







'ＬＩＮＱ的オペレーター -----------------------------------------------------------

'オペレーターのデリゲートは列挙時に Instantiate() され、列挙時の状態も保持する。

'　オペレータオブジェクトは、以下のメンバを持つ。
'　　・ SourceIterator  ... EnumVariant または 列挙ソースそのもの
'　　・ Expression      ... オペレータで使用する処理のデリゲート。
'　　・ Variable        ... 列挙中に変動し、列挙ごとに初期化される値。
'
'オペレータの Enumrator を生成する際、これらの値もセット・初期化する。





Public Function qSelect(selector_As_sourceIn_to_resultOut_) As Enumerator
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(selector_As_sourceIn_to_resultOut_)
    
'    If selector_ Is Nothing Then err.Raise'エラー処理
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    Set eop_.Expression = selector_
    
    
    Set qSelect = New Enumerator
    
    qSelect.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpSelect")
    
End Function




Public Function qWhere(predicate_As_sourceIn_to_isPassedOut_) As Enumerator
    
    Dim predicate_ As IFunc:    Set predicate_ = toFunc_(predicate_As_sourceIn_to_isPassedOut_)
    
'    If predicate_ Is Nothing Then err.Raise'エラー処理
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    Set eop_.Expression = predicate_
    
    
    Set qWhere = New Enumerator
    
    qWhere.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpWhere")
    
End Function




Public Function qTake(count_&) As Enumerator
    
'    if count_ < 0 then err.Raise 'エラー
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Variable = count_
    
    
    Set qTake = New Enumerator
    
    qTake.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpTake")
    
End Function




Public Function qSkip(count_&) As Enumerator
    
'    if count_ < 0 then err.Raise 'エラー
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Variable = count_
    
    
    Set qSkip = New Enumerator
    
    qSkip.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpSkip")
    
End Function




Public Function qSpan(count_&) As Enumerator
    
'    if count_ < 0 then err.Raise 'エラー
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Variable = count_
    
    
    Set qSpan = New Enumerator
    
    qSpan.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpSpan")
    
End Function




Public Function qSelectMany( _
 _
    Optional collectionSelector_As_sourceIn_to_collectionOut_, _
 _
    Optional resultSelector_As_collectionIn_elementIn_to_resultOut_ _
 _
) As Enumerator
    
    Dim colSelector_ As IFunc:  Set colSelector_ = toFunc_(collectionSelector_As_sourceIn_to_collectionOut_)
    
    Dim resSelector_ As IFunc:  Set resSelector_ = toFunc_(resultSelector_As_collectionIn_elementIn_to_resultOut_)
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Expression = Array(colSelector_, resSelector_)
    
    Dim innerIterator_  As IEnumVARIANT '内部イテレータ変数の初期化
    Dim elements_       As Object       '※現状は collectionSelector の結果はオブジェクトのみ、array は受け付けない
    
    eop_.Variable = Array(innerIterator_, elements_)
    
    
    Set qSelectMany = New Enumerator
    
    qSelectMany.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpSelectMany")
    
End Function





Public Function qGroupBy( _
 _
    Optional keySelector_As_sourceIn_to_keyOut_, _
 _
    Optional elementSelector_As_sourceIn_to_groupElementOut_, _
 _
    Optional resultSelector_As_keyIn_groupElementIn_to_resultOut_ _
 _
) As Enumerator
    
    Dim keySelector_ As IFunc:  Set keySelector_ = toFunc_(keySelector_As_sourceIn_to_keyOut_)
    
    Dim elmSelector_ As IFunc:  Set elmSelector_ = toFunc_(elementSelector_As_sourceIn_to_groupElementOut_)
    
    Dim resSelector_ As IFunc:  Set resSelector_ = toFunc_(resultSelector_As_keyIn_groupElementIn_to_resultOut_)
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Expression = Array(keySelector_, elmSelector_, resSelector_)
    
    Dim innerIterator_      As IEnumVARIANT '内部イテレータ変数の初期化
    Dim groupDictionary_    As Dictionary
    Dim key_
    
    eop_.Variable = Array(innerIterator_, groupDictionary_, key_)
    
    
    Set qGroupBy = New Enumerator
    
    qGroupBy.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpGroupBy")
    
End Function






Private Function qFlatAll(Optional resultSelector_) As Enumerator
    
    Set qFlatAll = New Enumerator
    
    Dim innerIterator_ As Enumerator: Set innerIterator_ = Nothing
    
    
End Function





'多次元配列の次元を１つ下げ、下位次元の列挙に変換する。出来るか？？

'　[ i, j, k, ... ] -> [ i ][ j, k, ... ] として扱い、i で回す。？？
'　→多次元配列は階層型列挙とみなして、平坦化でやるでもいいかも とすると多次元配列用イテレータかな chank 的なものでいいのかな

Private Function qDownDimension() As Enumerator
    
    
    
End Function

'意味ないかも？
'そもそも全ての要素の意味が同じならば（単なる座標的な意味でしかないなら）、フラットな列挙でよい。
'逆に例えばＣＳＶのテーブルやエクセルの表みたいなものだとしたら、それは構造的に扱う仕組みを作るべきだろう。
'　→それがめんどくさい場合、という需要もあるか？





'カウントごとにまとめて、外部イテレータを作成する。次に流れるのは外部イテレータで、エントリーは内部イテレータである。できる？？

Private Function qLot(count_&) As Enumerator
    
    Dim eop_ As New EnumOperatorProcs
    
    Set eop_.Variable = Ary.CNew.Alloc(count_)
    
End Function






'すべての要素を平坦化しつつ、親要素も含めて射影する。

Private Function qExprole(Optional relaySelector_, Optional terminalSelector_) As Enumerator
    
    
    
End Function









Public Function qScan(func_As_accumulateIn_sourceIn_to_accumulateOut_, Optional resultSelector_As_accumulateIn_to_resultOut_) As Enumerator
    
    Dim accFunc_        As IFunc:   Set accFunc_ = toFunc_(func_As_accumulateIn_sourceIn_to_accumulateOut_)
    
'    if accfunc_ is nothing then err.Raise 'エラー
    
    Dim resSelector_    As IFunc:   Set resSelector_ = toFunc_(resultSelector_As_accumulateIn_to_resultOut_)
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Expression = Array(accFunc_, resSelector_)
    
    Dim seed_
    
    eop_.Variable = seed_
    
    
    Set qScan = New Enumerator
    
    qScan.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpScanFirstStep")
    
    
End Function


Public Function qScanWithSeed(seed_, func_As_accumulateIn_sourceIn_to_accumulateOut_, Optional resultSelector_As_accumulateIn_to_resultOut_) As Enumerator
    
    Dim accFunc_        As IFunc:   Set accFunc_ = toFunc_(func_As_accumulateIn_sourceIn_to_accumulateOut_)
    
'    if accfunc_ is nothing then err.Raise 'エラー
    
    Dim resSelector_    As IFunc:   Set resSelector_ = toFunc_(resultSelector_As_accumulateIn_to_resultOut_)
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Expression = Array(accFunc_, resSelector_)
    
    If IsObject(seed_) _
        Then Set eop_.Variable = seed_ _
        Else eop_.Variable = seed_
    
    
    Set qScanWithSeed = New Enumerator
    
    qScanWithSeed.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpScanFirstStep2")
    
    
End Function







Public Function qRetrieveWhen(comparer_As_prevSourceIn_thisSourceIn_to_isEstablishedOut_, Optional selector_As_sourceIn_to_preConversionOut_) As Enumerator
    
    Dim comparer_ As IFunc:     Set comparer_ = toFunc_(comparer_As_prevSourceIn_thisSourceIn_to_isEstablishedOut_)
    
'    if comparer_ is nothing then err.Raise 'エラー処理
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(selector_As_sourceIn_to_preConversionOut_)
    
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Expression = Array(comparer_, selector_)
    
    Dim prevElement_
    
    eop_.Variable = prevElement_
    
    
    Set qRetrieveWhen = New Enumerator
    
    qRetrieveWhen.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpDistinctUntil")
    
End Function








Public Function qConcat(secondEnumerableSource_ As Object) As Enumerator
    
'    if secondEnumerableSource_ is nothing then err.Raise ’エラー処理
    
    
    Set qConcat = New Enumerator
    
    qConcat.SetOperator Me, makeUnionOperationDelegate_(secondEnumerableSource_, "OpUnionAll")
    
End Function

Public Function qUnion(secondEnumerableSource_ As Object) As Enumerator
    
'    if secondEnumerableSource_ is nothing then err.Raise ’エラー処理
    
    
    Set qUnion = New Enumerator
    
    qUnion.SetOperator Me, makeUnionOperationDelegate_(secondEnumerableSource_, "OpUnion")
    
End Function

Public Function qDistinct() As Enumerator
    
    Set qDistinct = New Enumerator
    
    qDistinct.SetOperator Me, makeUnionOperationDelegate_(Nothing, "OpUnion")
    
End Function


Private Function makeUnionOperationDelegate_(secondEnumerableSource_ As Object, operationName_$) As Delegate
    
    Dim eop_ As New EnumOperatorProcs
    
    Dim iterator_   As IEnumVARIANT
    Dim nextSource_ As Object:          Set nextSource_ = secondEnumerableSource_
    Dim distincts_  As Dictionary
    
    eop_.Variable = Array(iterator_, nextSource_, distincts_)
    
    Set makeUnionOperationDelegate_ = Delegate.CNew.SetMethod(eop_, operationName_)
    
End Function







Public Function qZip(secondEnumerableSource_ As Object, Optional resultSelector_As_firstIn_secondIn_to_resultOut_) As Enumerator
    
'    if secondEnumerableSource_ is nothing then err.Raise ’エラー処理
    
    Dim resSelector_    As IFunc:   Set resSelector_ = toFunc_(resultSelector_As_firstIn_secondIn_to_resultOut_)
    
    
    
    Dim eop_ As New EnumOperatorProcs
    
    Set eop_.Expression = resSelector_
    
    Dim iterator_       As IEnumVARIANT     ' secondSource_ と同じ領域を使用してもよいが、まぁわかりやすく別に。
    Dim secondSource_   As Object:          Set secondSource_ = secondEnumerableSource_
    
    eop_.Variable = Array(iterator_, secondSource_)
    
    
    Set qZip = New Enumerator
    
    qZip.SetOperator Me, Delegate.CNew.SetMethod(eop_, "OpZip")
    
End Function





' Expression デリゲート設定メソッド - - - - - - - - -

'オペレータをセットした後で、セレクタなどを設定できる。
'　オペレータに対応していない式を設定すると、エラーで通知する。


Public Function qSetCollectionSelector(collectionSelector_) As Enumerator
    Set qSetCollectionSelector = Me
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(collectionSelector_)
    
'    if selector_ is nothing then err.Raise 'エラー
    
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpSelectMany"):  i = 0
         
        Case Else: Exit Function '対応していないエラー
    
    End Select
    
    operation_.Target.SetExpression selector_, i
    
End Function


Public Function qSetResultSelector(resutlSelector_) As Enumerator
    Set qSetResultSelector = Me
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(resutlSelector_)
    
'    if selector_ is nothing then err.Raise 'エラー
    
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpSelectMany"):  i = 1
        Case operation_.HasMethod("OpGroupBy"):     i = 2
         
        Case Else: Exit Function '対応していないエラー
    
    End Select
    
    operation_.Target.SetExpression selector_, i
    
End Function


Public Function qSetKeySelector(keySelector_) As Enumerator
    Set qSetKeySelector = Me
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(keySelector_)
    
'    if selector_ is nothing then err.Raise 'エラー
    
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpGroupBy"):  i = 0
         
        Case Else: Exit Function '対応していないエラー
    
    End Select
    
    operation_.Target.SetExpression selector_, i
    
End Function


Public Function qSetElementSelector(elementSelector_) As Enumerator
    Set qSetElementSelector = Me
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(elementSelector_)
    
'    if selector_ is nothing then err.Raise 'エラー
    
    
    Dim i&
    Select Case True
        
        Case operation_.HasMethod("OpGroupBy"):  i = 1
         
        Case Else: Exit Function '対応していないエラー
        
    End Select
    
    operation_.Target.SetExpression selector_, i
    
End Function


' - - - - - - - - - - - - -



' ---------------------------------------------------------------------------------







'変換 -----------------------------------------------------------------------------



'コレクション化

'　元がコレクションでも、コピーする。

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






'配列化
'　領域確保に個数の取得が必要なため、一旦全要素をコレクション化してから配列のコピー元としている。

Public Function ToAry(Optional baseIndex_&) As Ary
    
    Dim src_
    
    If HasCountableSource _
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






'グルーピング辞書を構築する。
'　qGroupBy() で列挙されるものと同じ。こちらは列挙せずに辞書を返すパターン。

' keySelector       （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> key : variant

' elementSelector   （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> element_result : variant

Public Function ToLookup(Optional keySelector_As_sourceIn_to_keyOut_, Optional elementSelector_As_sourceIn_to_groupElementOut_) As Dictionary
    
    Dim keySelector_ As IFunc:  Set keySelector_ = toFunc_(keySelector_As_sourceIn_to_keyOut_)
    
    Dim elmSelector_ As IFunc:  Set elmSelector_ = toFunc_(elementSelector_As_sourceIn_to_groupElementOut_)
    
    
    Set ToLookup = makeGroupDictionary_(Me.NewEnum, keySelector_, elmSelector_)
    
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
        
        
        
        'キーが匿名オブジェクトなら、全要素を文字列化してシリアライズする。
        '　匿名オブジェクトは、
        
        Dim stringKey_$
        
        stringKey_ = Grouping.SeriarizeIfAnonymous(key_)
        
        
        
        '要素をグループコレクションに追加する。
        
        getGrouping_(groups_, key_, stringKey_).Elements.Add element_
        
        
    Loop
    
    
    Set makeGroupDictionary_ = groups_
    
End Function


'要素をグループコレクションに追加する｡
'　辞書から Grouping を取り出す。登録されていなければ新規 Grouping を登録する。

Private Function getGrouping_(groups_ As Dictionary, key_, stringKey_$) As Grouping
    
    
    ' key_ が辞書に登録されていればそれを返す。
    
    On Error Resume Next
    
    If stringKey_ = "" _
        Then Set getGrouping_ = groups_(key_) _
        Else Set getGrouping_ = groups_(stringKey_)
    
    If Not getGrouping_ Is Nothing Then Exit Function
    
    
    '登録されていなければ新規作成、登録、そして返す。
    
    Set getGrouping_ = New Grouping
    
    If IsObject(key_) _
        Then Set getGrouping_.Key = key_ _
        Else getGrouping_.Key = key_
    
    
    If stringKey_ = "" _
        Then Set groups_(key_) = getGrouping_ _
        Else Set groups_(stringKey_) = getGrouping_
    
End Function







'列挙から辞書を構築する。

' keySelector       （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> key : variant

' elementSelector   （省略時は、流れてくる値をそのまま次へ流す）
'　element : variant =>> element_result : variant

Public Function ToDictionary(Optional keySelector_As_sourceIn_to_keyOut_, Optional elementSelector_As_sourceIn_to_groupElementOut_) As Dictionary
    
    Dim keySelector_ As IFunc:  Set keySelector_ = toFunc_(keySelector_As_sourceIn_to_keyOut_)
    
    Dim elmSelector_ As IFunc:  Set elmSelector_ = toFunc_(elementSelector_As_sourceIn_to_groupElementOut_)
    
    
    Set ToDictionary = makeDictionary_(Me.NewEnum, keySelector_, elmSelector_)
    
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
        
        
        
        '要素を辞書に追加する。
        'キーが匿名オブジェクトなら、全要素を文字列化してシリアライズする。
        
        dict_.Add Grouping.ToAyKey(key_), element_
        
        
    Loop
    
    
    Set makeDictionary_ = dict_
    
End Function







' ---------------------------------------------------------------------------------








'結果取得系 ---------------------------------------------------------------------------------



'要素の個数を返す。ソースがカウント可能なら、直接のカウントを返す。

' predicate_
'　element : variant =>> is_passed : boolean

Public Property Get Count(Optional predicate_As_sourceIn_to_isPassedOut_) As Long
    
    Dim predicate_ As IFunc:    Set predicate_ = toFunc_(predicate_As_sourceIn_to_isPassedOut_)
    
    If IsBlank Then Exit Property
    
    
    Select Case True
        
        Case Not predicate_ Is Nothing, enumerableSource_ Is Nothing, TypeOf enumerableSource_ Is Enumerator
            
            Dim i
            For Each i In Me
                If predicate_.xExec01(i) Then Count = Count + 1
            Next
            
            
        Case HasCountableSource
            
            Count = asObject_(enumerableSource_).Count
            
            
    End Select
    
End Property





Public Function Aggregate(func_As_accumulateIn_sourceIn_to_accumulateOut_, Optional resultSelector_As_accumulateIn_to_resultOut_) As Variant
    
    Dim i
    For Each i In qScan(func_As_accumulateIn_sourceIn_to_accumulateOut_)
        VariantCopyInd Aggregate, i
    Next
    
    Dim resSelector_ As IFunc
    Set resSelector_ = toFunc_(resultSelector_As_accumulateIn_to_resultOut_)
    
    If Not resSelector_ Is Nothing Then
        
        VariantCopyInd Aggregate, resSelector_.xExec01(Aggregate)
        
    End If
    
End Function

Public Function AggregateWithSeed(seed_, func_As_accumulateIn_sourceIn_to_accumulateOut_, Optional resultSelector_As_accumulateIn_to_resultOut_) As Variant
    
    Dim i
    For Each AggregateWithSeed In qScanWithSeed(seed_, func_As_accumulateIn_sourceIn_to_accumulateOut_)
        VariantCopyInd AggregateWithSeed, i
    Next
    
    Dim resSelector_ As IFunc
    Set resSelector_ = toFunc_(resultSelector_As_accumulateIn_to_resultOut_)
    
    If Not resSelector_ Is Nothing Then
        
        VariantCopyInd AggregateWithSeed, resSelector_.xExec01(AggregateWithSeed)
        
    End If
    
End Function





'要素の値を合計する。要素が数値以外なら無視する。
'　null が一つでも入っていた時の結果については、本家は null になるがどうしようか、ＶＢＡの暗黙に準拠する感じでいいだろうか

Public Function Sum(Optional selector_As_sourceIn_to_conversionOut_) As Variant
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(selector_As_sourceIn_to_conversionOut_)
    
    
    Sum = 0& 'とりあえず数値だということを示し、文字列連結にならないように備える。
    
    Dim i
    For Each i In Me
        
        If Not selector_ Is Nothing Then i = selector_.xExec01(i)
        
        If IsNumeric(i) Then Sum = Sum + i
        
    Next
    
End Function



' Max / Min
'　流れてくる値の変換処理等は特にされない。
'　オブジェクトはオブジェクトのまま比較され（デフォルトプロパティに依存する）、結果もオブジェクトが返る。
'　比較できない場合も、特にエラー処理はされない。

' selector　比較前に実行される
'　source : variant => conversion : variant

Public Function Max(Optional selector_As_sourceIn_to_conversionOut_) As Variant
    
'    Max = choose_(Delegate.CNew(EnumOperatorProcs, "IsMax"), selector_As_sourceIn_to_conversionOut_)
    
    Dim i
    For Each i In qRetrieveWhen(Delegate.CNew(EnumOperatorProcs, "IsMax"), selector_As_sourceIn_to_conversionOut_)
        VariantCopyInd Max, i
    Next
    
End Function

Public Function Min(Optional selector_As_sourceIn_to_conversionOut_) As Variant
    
'    Min = choose_(Delegate.CNew(EnumOperatorProcs, "IsMin"), selector_As_sourceIn_to_conversionOut_)
    
    Dim i
    For Each i In qRetrieveWhen(Delegate.CNew(EnumOperatorProcs, "IsMin"), selector_As_sourceIn_to_conversionOut_)
        VariantCopyInd Min, i
    Next
    
End Function


'前の保存値と今回の値を比較関数に通して、真を返した場合保存値を入れ替える。

Private Function choose_(comparer_As_prevSourceIn_thisSourceIn_to_isEstablishedOut_, Optional selector_As_sourceIn_to_conversionOut_) As Variant
    
    Dim comparer_ As IFunc:     Set comparer_ = toFunc_(comparer_As_prevSourceIn_thisSourceIn_to_isEstablishedOut_)
    
'    if comparer_ is nothing then err.Raise 'エラー処理
    
    Dim selector_ As IFunc:     Set selector_ = toFunc_(selector_As_sourceIn_to_conversionOut_)
    
    
    Dim e_ As IUnknown:  Set e_ = Me.NewEnum
    
    If Not xCom.CallNext_EnumVariant(e_, choose_) Then Exit Function
    
    
    Dim candidate_
    
    Do While xCom.CallNext_EnumVariant(e_, candidate_)
        
        If Not selector_ Is Nothing Then VariantCopy candidate_, selector_.xExec01(candidate_)
        
        If comparer_.xExec02(choose_, candidate_) Then VariantCopy choose_, candidate_
        
    Loop
    
End Function



' ---------------------------------------------------------------------------------







'列挙 ---------------------------------------------------------------------------------


Public Function ForEach(expression_) As Enumerator
    Set ForEach = Me
    
    Dim exp_ As IFunc
    Set exp_ = toFunc_(expression_)
    
'    If exp_ Is Nothing Then err.raise'エラー処理
    
    Dim i
    For Each i In Me
        
        exp_.xExec01 i
        
    Next
    
End Function


' ---------------------------------------------------------------------------------






'その他 ---------------------------------------------------------------------------------




'ソースがセットされていなければ真を返す。

Public Property Get IsBlank() As Boolean
    
    IsBlank = enumerableSource_ Is Nothing And operation_ Is Nothing
    
End Property




'ソースがカウント可能（.Countプロパティを持つ）ならば真を返す。ただし、Enumerator は除く。

Public Function HasCountableSource() As Boolean
    
    On Error Resume Next
    
    If TypeOf enumerableSource_ Is Enumerator Then Exit Function
    
    HasCountableSource = IsNumeric(asObject_(enumerableSource_).Count)
    
End Function




' ---------------------------------------------------------------------------------











' 内部用 =============================================================================================




'ひな形 EnumOperatorProcs.xxx デリゲートをインスタンス複製し、ソースオブジェクト（ EnumVariant か列挙可能なオブジェクト）をセットして返す。

Private Function instantiateOperator_(operation_ As Delegate, srcEnum_ As IUnknown) As Delegate
    
    Dim newEop_ As EnumOperatorProcs
    
    Set newEop_ = operation_.Target.Instantiate
    
    
    Dim newDelegate_ As Delegate
    
    Set newDelegate_ = operation_.Clone
    
    
    Set newEop_.SourceIterator = srcEnum_
    
    Set instantiateOperator_ = newDelegate_.SetMethod(newEop_)
    
End Function





'キャスト　unknown → object

Private Function asObject_(unk_ As IUnknown) As Object
    
    Set asObject_ = unk_
    
End Function





'文字列コードなら Delegate にして返す。アドレスの可能性があれば FuncPointer にして返す。
'　Delegate などの IFunc ならそのまま返す。missing なら Nothing を返す。

Private Function toFunc_(expression_) As IFunc
    
    If IsMissing(expression_) Then Exit Function
    
    Select Case True
        
        Case TypeOf expression_ Is IFunc:                   Set toFunc_ = expression_
        
        Case varType(expression_) = vbString:               Set toFunc_ = Capt(CStr(expression_))
        
        Case varType(expression_) = varType(cNullPointer):  Set toFunc_ = FuncPointer.CNew(CLngPtr(expression_))
        
    End Select
    
End Function









