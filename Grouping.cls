VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Grouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







' Grouping インスタンスメンバ ============================================================




Public Key          As Variant

Private elements_   As New Collection







Public Property Get Elements() As Collection
Attribute Elements.VB_UserMemId = 0
    
    Set Elements = elements_
    
End Property


Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    
    If elements_ Is Nothing Then Exit Function
    
    Set NewEnum = elements_.[_NewEnum]
    
End Function








'非インスタンス関数 ==============================================


' Grouping に直接関係ない関数を持たせておくのはよくないかも…
'　でもどこに置くか…
'結構ぎちぎちになってきててこんがらがってきた…



'匿名オブジェクト関係 ------------------------------------------------




' Collection/Dictionary を匿名オブジェクトとして扱える。
'　第一要素のキー名を "ay"、値を Enumerable.AnonymousIdentify（オブジェクト参照値）とする。

' .net の Dictionary の key で匿名型は、
'　・参照アドレスではなく内部の値がキーになる。
'　・内側の匿名型もキーになる。通常のオブジェクトは内側にあっても参照アドレスがキーになるもよう。



'匿名オブジェクトを識別するために、匿名オブジェクトの先頭に持たせる ay 識別子。

Public Property Get AnonymousIdentify() As Collection
    
    Static anonymousIdentify_ As New Collection
    
    Set AnonymousIdentify = anonymousIdentify_
    
End Property



'オブジェクトをキーとする用途などのために、匿名オブジェクトならシリアライズ（文字列化）して返し、そうでなければそのまま返す。

Public Function ToAyKey(key_) As Variant
    
    Dim stringKey_$:    stringKey_ = SeriarizeIfAnonymous(key_)
    
    Select Case True
        
        Case stringKey_ <> "":  ToAyKey = stringKey_
        
        Case IsObject(key_):    Set ToAyKey = key_
        
        Case Else:              ToAyKey = key_
        
    End Select
    
End Function



'匿名オブジェクトの要素を文字列に変換して結合する。オブジェクトは「型名:アドレス」として文字列化する。
'　匿名オブジェクト以外が渡された場合は空文字を返す。

Public Function SeriarizeIfAnonymous(value_) As String
    
    
    If Not isAnonymousObject_(value_) Then Exit Function
    
    
    Dim arr_() As String
    ReDim arr_(value_.Count - 1 - 1) '匿名型判別キー "ay" の分を減じる
    
    Dim element_, i&
    For Each element_ In items_(value_)
        
        If Not isIdentify_(element_) Then
            
            Dim seriarized_$:   seriarized_ = SeriarizeIfAnonymous(element_)
            
            If seriarized_ <> "" Then element_ = seriarized_
            
            If IsObject(element_) _
                Then arr_(i) = TypeName(element_) & ":" & CStr(ObjPtr(element_)) _
                Else arr_(i) = CStr(element_)
            
            i = i + 1
            
        End If
        
    Next
    
    SeriarizeIfAnonymous = "{ '" & Join(arr_, "', '") & "' }"
    
End Function





' value_ が匿名オブジェクトなら真を返す。それ以外のすべてなら偽を返す。

Private Function isAnonymousObject_(value_) As Boolean
    
    On Error Resume Next
    
    isAnonymousObject_ = (value_("ay") Is AnonymousIdentify) 'キー"ay"に ay 識別子が格納されていれば匿名オブジェクトとする。
    
End Function



' element_ が ay 識別子なら真を返す。

Private Function isIdentify_(element_) As Boolean
    
    On Error Resume Next
    
    isIdentify_ = (element_ Is AnonymousIdentify)
    
End Function



' For Each 列挙のために、source_ が辞書なら .Items を、それ以外なら source_ をそのまま返す。

Private Function items_(source_) As Variant
    
    If TypeOf source_ Is Dictionary Then
        
        Dim d_ As Dictionary:   Set d_ = source_
        
        items_ = d_.Items
        
        Exit Function
        
    End If
    
    Set items_ = source_
    
End Function



' --------------------------------------------------------------------------------












