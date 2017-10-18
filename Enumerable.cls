VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







' enumerableSource_ が対象外の場合は、空のソースを所持する Enumerator が返る。

Public Function From(enumerableSource_) As Enumerator
Attribute From.VB_UserMemId = 0
    
    Select Case True
        
        
        Case TypeOf enumerableSource_ Is Enumerator
            
            '素通し
            
            Set From = enumerableSource_
            
            
            
        Case TypeOf enumerableSource_ Is Ary
            
            ' Ary は列挙時に EnumVariant を返せないので、外部からイテレータをデリゲートとして設定してやる。
            
            Set From = New Enumerator
            
            Dim srcAry_ As Ary:     Set srcAry_ = enumerableSource_
            
            From.SetIterator srcAry_, Delegate.CNew.SetMethod(New EnumOperatorProcs, "OpIteratorAry")
            
            
            
        Case TypeOf enumerableSource_ Is Object , TypeOf enumerableSource_ Is IEnumVARIANT
            
            '普通のオブジェクトは列挙時にＣＯＭオブジェクトとしての EnumVarint を取得するため、オブジェクトだけ渡せばいい。
            
            Set From = New Enumerator
            
            Dim srcUnk_ As IUnknown:  Set srcUnk_ = enumerableSource_
            
            From.SetSource srcUnk_
            
            
            
        Case IsArray(enumerableSource_)
            
            ' 配列は値コピーとなってしまう
            
            Set From = Enumerable(Ary.CopyFrom(enumerableSource_))
            
            
            
        Case Else
            
            'ブランクを返す。
            
            Set From = New Enumerator
            
            
    End Select
    
End Function





'一定範囲の Enumerable( Ary ) を返す

Public Function Range(start_#, count_&, Optional span_# = 1) As Enumerator
    
    Dim arr_ As New Ary
'    Dim arr_ As New Collection
    
    arr_.Alloc count_, vbLong
    
    Dim i&, num_&
    For num_ = start_ To start_ - 1 + count_ Step span_
        
        arr_(i) = num_
'        arr_.Add num_
        
        i = i + 1
    Next
    
    Set Range = Enumerable(arr_)
    
End Function


Private Function Range2(start_#, last_#, Optional freq_&) As Enumerator
    
    Dim arr_ As New Ary
'    Dim arr_ As New Collection
    
    
    Dim range_#:    range_ = last_ - start_
    
    Dim span_#:     span_ = range_ / IIf(freq_ > 0, freq_, range_)
    
    
    arr_.Alloc Fix(range_ / span_) + Sgn(range_ / span_), vbDouble
    
    Dim i&, num_&
    For num_ = start_ To last_ Step span_
        
        arr_(i) = num_: Debug.Print num_
'        arr_.Add num_
        
        i = i + 1
    Next
    
'    if
    
'    Set Range = Enumerable(arr_)
    
End Function











