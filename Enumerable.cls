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







' enumerableSource_ が列挙対象外の場合は、空のソースを所持する Enumerator が返る。

Public Function From(enumerableSource_) As Enumerator
Attribute From.VB_UserMemId = 0
    
    Select Case True
        
        
        Case Not IsObject(enumerableSource_) And IsArray(enumerableSource_)
            
            ' 配列は値コピーとなってしまう
            
            Set From = Enumerable(Ary.CopyFrom(enumerableSource_))
            
            
        
        Case enumerableSource_ Is Nothing
            
            'ブランクを返す。
            
            Set From = New Enumerator
            
            
            
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
            
            
            
        Case Else
            
            '対応できない場合
            
'            Err.Raise 'エラー処理
            
            
    End Select
    
End Function





'一定範囲を列挙するイテレーターを返す

Public Function Range(start_#, count_&, Optional span_# = 1) As Enumerator
    
    Set Range = New Enumerator
    
    Dim eop_ As New EnumOperatorProcs
    
    eop_.Variable = Array(start_, span_, count_)
    
    Range.SetIterator Nothing, Delegate.CNew.SetMethod(eop_, "OpGeneratorRange")
    
End Function





'配列の次元を指定分だけまとめる列挙子を生成する。

Private Function qLot(sourceAry_ As Ary, applyDimensions_&) As Enumerator
    
    
    
End Function








