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





'構造体定義 ============================================================



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




'変則的 Variant 内部レイアウト

Private Type VariantStructWithCollection
    vt              As Integer
    reserve0        As Integer
    serialLength    As Long 'reserve1    As Integer 'reserve2    As Integer
    '次元を無視した通し個数。Variant の予約部分を使用する。一抹の不安はある笑
    pEntity0        As LongPtr
    dimDefs         As Collection 'pEntity1 As LongPtr
    'メモリ配置としてはオブジェクトはポインターと同等なので、Nothing は cNullPointer と同等
End Type







'メンバ宣言 ============================================================





'配列保持

Private arr_ As VariantStructWithCollection ' Variant 型変数に MoveMemory できるようにこのメモリレイアウトで保持している。







'関数定義 ============================================================




'オブジェクトインスタンスの新規作成

Public Function CNew() As Ary
    
    Set CNew = New Ary
    
End Function





'配列確保 ------------------------------------------------


'一次元配列 - - - - -


'一次元配列を確保（長さを指定）

Public Function Alloc(length_&, Optional varType_ As VbVarType = vbVariant) As Ary
    Set Alloc = Me
    
    Free
    
    
    arr_.vt = varType_ Or cVarArray
    
    arr_.pEntity0 = SafeArrayCreateVector(varType_, 0&, length_)
    
    arr_.serialLength = length_
    
End Function


'一次元配列確保（先頭添え字と長さを指定）

Public Function AllocBound(baseIndex_&, length_&, Optional varType_ As VbVarType = vbVariant) As Ary
    Set AllocBound = Me
    
    Free
    
    
    arr_.vt = varType_ Or cVarArray
    
    arr_.pEntity0 = SafeArrayCreateVector(varType_, baseIndex_, length_)
    
    arr_.serialLength = length_
    
End Function



'多次元配列 - - - - -


'多次元配列１要素の長さを宣言する。配列確保するのにコレクション作るとか本末転倒感…。

Public Function DimDef(length_&) As Ary
    Set DimDef = Me
    
    Free
    
    If arr_.dimDefs Is Nothing Then Set arr_.dimDefs = New Collection
    
    
    arr_.dimDefs.Add Array(0&, length_)
    
    
End Function


'多次元配列１要素の先頭添え字と長さを宣言する。

Public Function DimDefBound(baseIndex_&, length_&) As Ary
    Set DimDefBound = Me
    
    Free
    
    If arr_.dimDefs Is Nothing Then Set arr_.dimDefs = New Collection
    
    
    arr_.dimDefs.Add Array(baseIndex_, length_)
    
    
End Function


'多次元要素の宣言をもとに、多次元配列を確保する。

Public Function DimAlloc(Optional varType_ As VbVarType = vbVariant) As Ary
    Set DimAlloc = Me
    
    If arr_.dimDefs Is Nothing Then Exit Function
    
    
    
    '多次元要素の宣言から safeArray 確保用構造体を生成する。
    
    Dim safeArrayBounds_() As SafeArrayBound:   ReDim safeArrayBounds_(arr_.dimDefs.Count - 1)
    
    Dim dimDefine_, i&, totalLength_&
    For Each dimDefine_ In arr_.dimDefs
        
        safeArrayBounds_(i).BaseIndex = dimDefine_(0)
        
        safeArrayBounds_(i).Elements = dimDefine_(1)
        
        totalLength_ = totalLength_ + dimDefine_(1)
        
        i = i + 1
    Next
    
    Set arr_.dimDefs = Nothing '宣言を開放
    
    
    ' safeArray を確保する。
    
    arr_.vt = varType_ Or cVarArray
    
    arr_.pEntity0 = SafeArrayCreate(varType_, i, safeArrayBounds_(0))
    
    arr_.serialLength = totalLength_
    
End Function


' -----------------------------------------------------------





'コピー／ムーブ ----------------------------


' Variant に格納された動的配列を Ary に移動する。
'　移動元には Empty が入る。

Public Function MoveFrom(ByRef ref_array_) As Ary
    Set MoveFrom = Me
    
    Free
    
    
    '移動先へコピーする。
    
    MoveMemory arr_, ref_array_, cSizeOfVariant
    
    If arr_.vt And cVarRef Then Exit Function '引数の元が Variant 以外で参照になっている場合など　エラー処理すべし
    
    'arr_.serialLength =
    
    
    '移動元を Empty にし、そちら側で開放されないようにする。
    
    MoveMemory ref_array_, Empty, cSizeOfVariant
    
    
End Function


'配列をコピーする。

Public Function CopyFrom(array_) As Ary
    Set CopyFrom = Me
    
    Free
    
    
    VariantCopy arr_, array_
    
    'arr_.serialLength =
    
End Function


' ------------------------------------------------------------






'解放 --------------------------------------------------------


'配列の解放

Public Function Free() As Ary
    Set Free = Me
    
    If arr_.pEntity0 = cNullPointer Then Exit Function
    
    
    '配列の破棄
    
    SafeArrayDestroy ByVal arr_.pEntity0
    
    
    ' SafeArray のクリア
    
    Dim blankArr_ As VariantStructWithCollection
    
    arr_ = blankArr_ '要素宣言も解放される
    
    
End Function

Private Sub Class_Terminate()
    
    Free
    
End Sub

' ---------------------------------------------------




'情報取得 -------------------------------------------


'次元ごとの要素個数を返す。次元を省略した場合、0 次元を指定したとみなす。

Public Property Get Length(Optional dimension_&) As Long
    
    SafeArrayGetUBound ByVal arr_.pEntity0, dimension_ + 1, out_ubound_:=Length
    
    Length = Length + 1
    
End Property

Public Property Get BaseIndex(Optional dimension_&) As Long
    
    SafeArrayGetLBound ByVal arr_.pEntity0, dimension_ + 1, out_lbound_:=BaseIndex
    
End Property


'次元の数を返す。

Public Property Get Dimension() As Long
    
    Dimension = SafeArrayGetDim(ByVal arr_.pEntity0)
    
End Property


'全要素数を返す。

Public Property Get Count() As Long
    
    Count = arr_.serialLength
    
End Property


' --------------------------------------------------------






'要素へのアクセス --------------------------------------------

'ary_.Item(0) のように使用することを想定している。

'配列の参照を返し、要素アクセスに配列コピーが生じないようにしている。
'　ただし、otherArr_ = ary_.Item のように外部のバリアント変数などに渡してはいけない。
'　Ary クラスが破棄されると配列本体も破棄され、そのあと参照を通してアクセスするとエクセルが落ちる。どうしたもんかね…。

Public Property Get Item() As Variant
Attribute Item.VB_UserMemId = 0
    
    
    ' SafeArray 参照 Varriant 構造体の構築
    
    Dim v_ As VariantStructWithCollection
    
    v_.vt = arr_.vt Or cVarRef
    
    v_.pEntity0 = VarPtr(arr_.pEntity0)
    
    
    '参照構造体を返す。
    
    MoveMemory Item, v_, cSizeOfVariant
    
    
End Property


'次元を無視した通しインデックスを指定して要素を取得する。取得できた場合は真、できなかった場合は偽を返す。

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
    
'    Set NewEnum = Enumerable(Me).NewEnum と同じ
    
    Dim op_ As New EnumOperatorProcs
    
    Set op_.SourceIterator = Me
    
    Set NewEnum = xCom.CreateEnumVariant(Delegate.CNew(op_, "OpIteratorAry"))
    
End Function





