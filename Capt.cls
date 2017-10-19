VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Capt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary


'ラムダ式のようなものを実現するクラス
'　動的に関数を生成し、以降はコードの文字列から辞書引きで取得できるようにする。
'　また取得した関数を Delegate オブジェクトにセットして、繰り返し実行できる。



'コード式文字列の書式

'　[ [ [引数リスト] : キャプチャ引数リスト ] => ] 行 ( 文 | <result> = 戻り値の式 | 空文 ) { ; 行 }
'
'　（特殊なケースとして単一行かつ文頭に「>」がある場合は、戻り値の式として解釈される）
'
'　例：
'　　x => x = x + 1                 ... x は ByRef であり、戻り値は Empty
'　　<result> = 2                   ... 引数を取らず 2 が返る
'　　> 2                            ... <result> = 2 と同じ
'　　x =>> x + 1                    ...「=>」と「>」は続けて書く必要があるため「=>>」となる（間にスペースが入ると行頭として解釈されなくなる）
'　　x : cap => x = x + cap         ... : の後はキャプチャ変数（ Capt.Capture() に渡した順に対応する）
'　　: cap1 =>> capt1 + 1           ... キャプチャ変数の前の : は省略できない
'　　x, y => x = x + 1; y = y + 1   ... 複数の文は ; で区切る（文字列以外の ; は全て改行に置き換わる）
'
'　複雑な例：
'　　x, y, z : capt1, capt2 => y = 1; z = x + y; <result> = x + 1;


'　・文字列は '' で括る。
'　　文字列中に「\'」と入力した場合、内部で一旦 escChar_ となり、最終的には「'」に展開される。
'　　（ escChar_ = Chr(1) だが、大丈夫だろうか？？）
'　　$'abc{def}ghi' とすると、def は変数名として展開される。その場合、「{{」「}}」が「{」「}」のエスケープとなる。
'
'　・単一行かつ、命令文行頭が「>」で始まるコードは、値を返す式として解釈される。
'　　例：   Capt( "> 0 + 1" )       ... 命令文は「> 0 + 1」となる。
'　　       Capt( "x =>> x + 1" )   ...「x 」「=>」「> x + 1」と解釈され、命令文は「> x + 1」となる。
'　　       Capt( "x =>> x + 1;" )  ...「x 」「=>」「> x + 1」「」と解釈され、命令文は「> x + 1」および「」（空行）の
'                                       ２行となるため、エラーとなる。
'　　※「>」必要なしの自動判別としたかったが、字句解析が面倒なのでやめました…。
'
'　・改行コードを含んでも問題ないが、内部で「;」に置き換えられる。
'
'　・引数／キャプチャ引数には型も指定できるが、呼び出し時に variant を経由して渡されるため、そこで型チェックに引っかかるため、実用できない。
'


'コード式文字列を辞書キーとして、動的に生成されたメソッドのデリゲートを呼び出す。

'辞書のキーとするコードは、素のままとした。
'　両端のトリムや小文字化、重複スペース除去等も考えたが、重複登録よりも関数名取得の軽さを優先することにした。
'　暫定要検討。

'本モジュール中の文字列比較は基本的にバイナリ比較
'　・レターよりも記号との比較が多いため
'　・文字列コード式の辞書引きもバイナリ比較（動的関数ゴミがたまることよりも辞書引きの速度を優先、意味ないか？）
'
'　以下の二か所でのみテキスト比較使用
'　　parseCode_()               -> replace()    ... "<result>" 置換
'　　buildCaptureVariables_()   -> split()      ... "as" 分割
'
'　本モジュールで使用している比較関数等
'　　replace(), split(), like



'正規表現使うべきだったかなぁ…。
'　参照設定が面倒と思ってやめたんだけど…スクリプティング使うし一緒だったかなぁ…。
'　そのうち置き換えていこう…。





'定数定義 ----

Private Const cMainCodeHolderName$ = "DynamicCode"
Private Const cTmpNameHeader$ = "ZTmp"
Private Const cTmpStdModuleName$ = "ZTmpDynamic"
Private Const cFuncNameHeader$ = "Func"
Private Const cCollNameHeader$ = "coll"
Private Const cDictNameHeader$ = "dict"
Private Const cNameForDynamicCodeSearch$ = "????#*"




'メンバ宣言 ============================================================





'文字列コード式 →  への対応辞書。

Private delegates_ As New Dictionary








Private escChar_$ '改行や文字列内「\'」のエスケープに使用する。

Private Sub initVariable_()
    
    delegates_.CompareMode = vbBinaryCompare
    
    escChar_ = Chr(1)
    
End Sub







'文字列コード式のキャプチャ ===========================================================


'コード式文字列とキャプチャパラメータを取り込む。デフォルトメソッド。

Public Function Capture(code_$, ParamArray captureParams_()) As IFunc
Attribute Capture.VB_UserMemId = 0
    
    If code_ = "" Then Exit Function '何も渡されない場合は Nothing を返す。→エラー処理にすべきか？
    
    
    
    'デリゲートを作成する／作成済みを取得する。
    
    Set Capture = getFunc_(CStr(code_))
    
    
    
    'キャプチャ引数があれば、疑似クロージャデリゲートにして返す。
    
    If IsMissing(captureParams_) Then Exit Function
    
    Dim d_ As Delegate
    Set d_ = Capture
    Set d_ = d_.Clone.SetMethod(object_:=d_.Target.Instantiate)
    
    d_.Target.Captures = captureParams_
    
    Set Capture = d_
    
    
End Function


'受け取ったコード式文字列から、実行すべき関数オブジェクトを返す。
'　生成済みならそれを返すが、未生成ならＶＢＡコードを動的生成する。

Private Function getFunc_(code_$) As Delegate
    
    
    On Error Resume Next
    
    Set getFunc_ = delegates_(code_)
    
    On Error GoTo 0
    
    
    If getFunc_ Is Nothing Then
        
        '未生成時は動的生成し、辞書に登録する。
        
        Set getFunc_ = createFunc_(code_)
        
        Set delegates_(code_) = getFunc_
        
    End If
    
End Function







'動的コードの生成 ==============================================================


'文字列コード式からＶＢＡクラスコードを動的生成する。
'　戻り値は関数オブジェクト。
'　生成したコードにはコメントとして、文字列コード式を残しておく（生成済み辞書構築時のガイドとなる）。
'　生成時は新規クラス（ cTempCodeClassName + n ）として生成され、CodeClean() 時に cMainCodeHolderName へまとめられる。
'　　※本当なら、新規生成時には動的コードホルダに直接追加したいのだが、
'　　　一度でもクラスがインスタンス化されると動的コード生成が反映されなくなるため、生成時は個別に生成せざるをえない。
'　　　標準モジュールなら即反映されるのだが、モジュールが書き換えられると他の関数アドレスも変わってしまうため、個別モジュールは必要。

Private Function createFunc_(code_$) As Delegate
    
    
    Dim funcName_$:     funcName_ = makeUniqueName_(cFuncNameHeader)                '関数名の生成
    
    Dim comment_$:      comment_ = "'" & Replace(code_, vbCrLf, escChar_) & "'"  '後ろの「'」は空欄消去防止用
    
    
    'コード式文字列をパースしてＶＢＡコードにする。
    
    Dim funcString_$:   funcString_ = parseCode_(funcName_, code_, comment_)
    
    
    
    '動的関数を新規クラスモジュールに書き出し、そのクラスインスタンスを取得する。
    
    Dim newClassName_$: newClassName_ = makeUniqueRndModuleName_(cTmpNameHeader)
    
    
    Dim newModule_ As CodeModule
    Set newModule_ = getCodeModule_(newClassName_, vbext_ct_ClassModule)
    
    newModule_.AddFromString makeDynamicCodeHead_(newClassName_)
    
    newModule_.AddFromString vbCrLf & funcString_
    
    
    Dim obj_ As Object
    Set obj_ = instantiate_(newClassName_, cTmpStdModuleName)
    
    Set createFunc_ = Delegate.CNew(obj_, funcName_)
    
End Function






'関数文字列を構築 ------------------------------------------------------


'渡されたコードを整形してモジュールへ格納できる関数にする。
 
Private Function parseCode_(funcName_$, code_$, comment_$) As String
    
    If Trim$(code_) = "" Then Exit Function
    
    
    
    Dim pureCode_$:     pureCode_ = code_
    
    
    Dim additionalFuncs_    As New Collection   '追加したサブ関数を文字列として登録する。
    
    Dim strInStrs_          As New Collection   '「文字列内文字列」保存用。
    
    
    
    
    '一時的に「文字列内文字列」を退避する。
    
    pureCode_ = Replace(pureCode_, "\'", escChar_)           '「\'」はエスケープする
    
    pureCode_ = escapeStrInStr_(pureCode_, "'", strInStrs_)     ' '' で括られたものを退避させる
    
    
    
    
    '匿名オブジェクト等の糖衣構文をパースする。
    
    pureCode_ = scanSugarSyntaxes_(pureCode_, additionalFuncs_)
    
    
    
    
    Dim args_$, capts_$, lineBlock_$
    
    
    'コード文字列を分解し、引数リスト・キャプチャリスト・文ブロックを生成する。
    
    dividFuncBlocks_ pureCode_, _
 _
        out_args_:=args_, out_capts_:=capts_, out_lineBlock_:=lineBlock_
    
    
    '関数文字列を構築する。
    
    parseCode_ = buildFuncString_(funcName_, args_, capts_, lineBlock_, comment_)
    
    
    
    
    
    '「文字列内文字列」を元に戻す。
    
    parseCode_ = unescapeStrInStr_(parseCode_, """", strInStrs_)    '退避されたものを "" で括って復元する
    
    parseCode_ = Replace(parseCode_, escChar_, "'")              '「\'」エスケープを「'」に直す
    
    
    
    
    '追加の関数を付加する。
    
    If additionalFuncs_.Count = 0 Then Exit Function
    
    parseCode_ = parseCode_ & vbCrLf & vbCrLf & joinCollection_(additionalFuncs_, vbCrLf & vbCrLf)
    
    
End Function





'コード文字列を分解し、引数リスト・キャプチャリスト・文ブロックを生成する。

Private Function dividFuncBlocks_(pureCode_$, ByRef out_args_$, ByRef out_capts_$, ByRef out_lineBlock_$)
    
    
    '引数ブロックと文ブロックに分割する。省略されていれば "" となる。
    
    Dim blocks_$():         blocks_ = Split(pureCode_, "=>")
    
    Dim paramBlock_$:       If UBound(blocks_) > 0 Then paramBlock_ = blocks_(0)    '引数／キャプチャリストを省略（「=>」ごと省略も含む）した場合は ""
    
    Dim lineBlock_$:        lineBlock_ = blocks_(UBound(blocks_))                   '省略されることはない（コード全体が "" となるが、その場合は関数作成されない）
    
    
    
    '引数ブロックを、引数リストとキャプチャリストに分割する。省略されていれば "" となる。
    
    Dim paramBlocks_$():    paramBlocks_ = Split(LTrim(paramBlock_), ":")           '引数ブロックが "" の場合は、ubound() が -1 になるぽい…。
    
    Dim args_$:             If UBound(paramBlocks_) > -1 Then args_ = paramBlocks_(0)
    
    Dim capts_$:            If UBound(paramBlocks_) > 0 Then capts_ = paramBlocks_(1)
    
    
    
    '「>」で始まる単一行の場合、その行を式と捉え、戻り値として返すようにする。
    
    If InStr(1, lineBlock_, ";") = 0 Then
        
        If lineBlock_ Like ">*" Then
            
            lineBlock_ = "VariantCopyInd <result>, " & Mid$(lineBlock_, 2)
            
        End If
        
    End If
    
    
    
    '戻り
    
    out_args_ = args_
    
    out_capts_ = capts_
    
    out_lineBlock_ = lineBlock_
    
End Function




' - - - - -


'関数文字列を構築する。

Private Function buildFuncString_(funcName_$, args_$, capts_$, lineBlock_$, comment_$) As String
    
    
    Dim captVarString_$:    captVarString_ = buildCaptureVariables_(capts_)
    
    Dim linesString_$:      linesString_ = buildCodeLines_(lineBlock_, funcName_)
    
    
    buildFuncString_ = _
 _
        "public function {funcName}( {a,b,c} ) as variant" & vbCrLf & _
 _
            comment_ & vbCrLf & _
 _
            formatCodeLine_(captVarString_) & _
 _
            formatCodeLine_(linesString_) & _
 _
        "end function"
    
    
    buildFuncString_ = Replace(buildFuncString_, "{funcName}", funcName_, 1, 1)
    buildFuncString_ = Replace(buildFuncString_, "{a,b,c}", args_, 1, 1)
    
End Function




'キャプチャ変数の宣言文を整形構築する。
'　（一応型を指定できるようにはしてあるが、引数リストが型未対応なので隠し機能ということで…）

Private Function buildCaptureVariables_(captureString_$) As String
    
    Dim captVariables_$():  captVariables_ = Split(captureString_, ",")
    
    
    Dim i&
    For i = 0 To UBound(captVariables_)
        
        
        Dim varName_$:  varName_ = Split(captVariables_(i), "as", , vbTextCompare)(0)
        
        
        captVariables_(i) = _
 _
            "dim " & captVariables_(i) & ":" & _
 _
            "VariantCopy " & varName_ & ", captures(" & i & ")"
        
        
    Next
    
    
    buildCaptureVariables_ = Join(captVariables_, vbCrLf & vbTab)
    
End Function




'コード本体を整形構築する。

Private Function buildCodeLines_(lineBlock_$, funcName_$) As String
    
    buildCodeLines_ = lineBlock_
    
    buildCodeLines_ = Replace(buildCodeLines_, "<result>", funcName_, , , vbTextCompare)    '戻り値の処理
    
    buildCodeLines_ = Replace(buildCodeLines_, ";", vbCrLf)                                 '「;」は全て改行化
    
    buildCodeLines_ = Replace(buildCodeLines_, vbCrLf, vbCrLf & vbTab)                      '簡易ながらインデントを付加
    
End Function




'１行を表す文字列にインデント（タブ）と改行を付加する。空文字には付加しない。

Private Function formatCodeLine_(codeLine_$) As String
    
    If codeLine_ = "" Then Exit Function
    
    formatCodeLine_ = vbTab & codeLine_ & vbCrLf
    
End Function



'------------------------------------------------------

















'汎用 =======================================================



'モジュール操作関連 -----------------------------------



'ＶＢＡのコードモジュールを取得する。
'　存在しない場合は cmType_ で指定されたモジュールを作成して返す。

Private Function getCodeModule_(cmName_$, cmType_ As vbext_ComponentType) As CodeModule
    
    On Error Resume Next
    
    
    Dim cmp_ As VBComponent
    Set cmp_ = ThisWorkbook.VBProject.VBComponents(cmName_)
    
    If cmp_ Is Nothing Then
        
        '見つからなければ作成して返す。
        
        Set getCodeModule_ = ThisWorkbook.VBProject.VBComponents.Add(cmType_).CodeModule
        
        getCodeModule_.Parent.Name = cmName_
        
        Exit Function
        
    End If
    
    
    If cmp_.Type <> cmType_ Then Exit Function  '取得したモジュールのタイプが異なる場合は Nothing が返る
    
    Set getCodeModule_ = cmp_.CodeModule
    
End Function




'該当するコードモジュールを全て返す。名前にはワイルドカードが使用できる。１つもなくても空のコレクションが返る。

Private Function getCodeModules_(nameWild_$, cmType_ As vbext_ComponentType) As Collection
    
    Dim targets_ As New Collection
    
    Dim vbc_ As VBComponent
    For Each vbc_ In ThisWorkbook.VBProject.VBComponents
        
        If vbc_.Type = vbext_ct_ClassModule Then
            
            If vbc_.Name Like nameWild_ Then _
 _
                targets_.Add vbc_.CodeModule
                
        End If
        
    Next
    
    Set getCodeModules_ = targets_
    
End Function




'モジュール内の関数名をリストアップする。

Private Function getFuncNames_(cm_ As CodeModule, Optional funcNameHeaderWild_$, Optional iLineStart_&) As Collection
    
    Dim funcNames_ As New Collection
    
    
    Dim iStart_&:   iStart_ = IIf(iLineStart_ = 0, cm_.CountOfDeclarationLines, iLineStart_)
    
    
    Dim i&
    For i = iStart_ To cm_.CountOfLines '一行ずつ検索
        
        
        Dim funcName_$:     funcName_ = cm_.ProcOfLine(i, vbext_pk_Proc)
        
        
        If funcName_ Like funcNameHeaderWild_ Then
            
            
            funcNames_.Add funcName_
            
            
            i = i + cm_.ProcCountLines(funcName_, vbext_pk_Proc) - 1
            
        End If
        
    Next
    
    
    Set getFuncNames_ = funcNames_
    
End Function







'新規クラスモジュールにコードを追加する。
'　CodeModule.AddFromString() の動作により、宣言セクションのすぐ後に挿入される。
'　ただし isOverWrite_ が真の場合は、全テキストが上書きされる。

Private Function addCodeToNewModule_(codeHolderName_$, codeString_$, Optional isOverWrite_ As Boolean) As CodeModule
    
    Dim newCm_ As CodeModule
    
    Set newCm_ = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_ClassModule).CodeModule
    
    
    newCm_.Name = codeHolderName_
    
    
    If Not isOverWrite_ Then
        
        newCm_.AddFromString vbCrLf & codeString_
        
    Else
        
        newCm_.DeleteLines 1, newCm_.CountOfLines
        
        newCm_.InsertLines 1, codeString_
        
    End If
    
    
    Set addCodeToNewModule_ = newCm_
    
End Function


'------------------------------------------------------




'文字列内文字列の一時置換 ------------------------------


'文字列の中の delim_ で括られた「文字列内文字列」を「""」（escapeDelim_×２）に置き換え、コレクションに退避する。
'　delim_ で split() した時の奇数部分が「文字列内文字列」にあたる。
'　strInStrs_ コレクションに、元の「文字列内文字列」を保存して返す。

Private Function escapeStrInStr_(code_$, delim_$, strInStrs_ As Collection, Optional escapeDelim_$ = """") As String
    
    Dim seps_$(): seps_ = Split(code_, delim_)
    
    Dim i&
    For i = 1 To UBound(seps_) Step 2
        
        strInStrs_.Add seps_(i)
        
        seps_(i) = ""
        
    Next
    
    escapeStrInStr_ = Join(seps_, escapeDelim_)
    
End Function




'文字列内文字列を復元する。
'　また、$'' 文字列の変数展開も行う。

Private Function unescapeStrInStr_(code_$, delim_$, strInStrs_ As Collection, Optional escapeDelim_$ = """") As String
    
    Dim seps_$(): seps_ = Split(code_, escapeDelim_)
    
    Dim i&, sis_ '$
    For Each sis_ In strInStrs_
        
        seps_(i + 1) = replaceVariable_(delim_, CStr(sis_), ref_prevSection_:=seps_(i))
        
        i = i + 2
        
    Next
    
    unescapeStrInStr_ = Join(seps_, delim_)
    
End Function




' $'abc{x}def' 中の {x} を変数に展開する。
'　「{」「}」は「{{」「}}」でエスケープされる。

'例：   $"abc{var_}efg"     -> "abc" & var_ & "efg"
'　　    "dsef{esefe}sef"   -> "dsef{esefe}sef"     ※「$」がないため展開されない
'　　   $"sec:{@3}"         -> "sec:" & @3          ※ＶＢＡの文法的に不正でもそのまま展開される
'　※ 例の delim_ は「"」とした。

Private Function replaceVariable_(delim_$, str_$, ByRef ref_prevSection_$) As String
    
    replaceVariable_ = str_
    
    If Not ref_prevSection_ Like "*$" Then Exit Function
    
    
    ref_prevSection_ = Mid$(ref_prevSection_, 1, Len(ref_prevSection_) - 1) '前節末尾の $ を省く
    
    
    replaceVariable_ = Replace(replaceVariable_, "{{", escChar_)
    
    replaceVariable_ = Replace(replaceVariable_, "{", delim_ & " & (")
    
    replaceVariable_ = Replace(replaceVariable_, escChar_, "{")
    
    
    replaceVariable_ = Replace(replaceVariable_, "}}", escChar_)
    
    replaceVariable_ = Replace(replaceVariable_, "}", ") & " & delim_)
    
    replaceVariable_ = Replace(replaceVariable_, escChar_, "}")
    
    
    replaceVariable_ = Replace(replaceVariable_, " & " & delim_ & delim_, "")
    
    replaceVariable_ = Replace(replaceVariable_, delim_ & delim_ & " & ", "")
    
End Function



' ----------------------------------------------------











'その他 ------------------------------------------------




'名前で指定された新規クラスをインスタンス化する。関数の動的生成を使用している。

Private Function instantiate_(className_$, stdModuleName_$) As Object
    
    'クラスインスタンス化コードを動的に生成（他に動的クラス生成の方法ないのか…？）
    
    Dim tmpCm_ As CodeModule
    
    Set tmpCm_ = getCodeModule_(stdModuleName_, vbext_ct_StdModule)
    
    Dim funcString_$:   funcString_ = "public function New_{className}() as {className}: set New_{className} = new {className}: end function"
    
    tmpCm_.AddFromString Replace(funcString_, "{className}", className_)
    
    
    'コードクラスをインスタンス化する。
    
    Set instantiate_ = Application.Run(stdModuleName_ & ".New_" & className_)
    
End Function







'使用されていない、ランダムなモジュール名を生成する。

Private Function makeUniqueRndModuleName_(namePrefix_$) As String
    
    On Error Resume Next
    
    
    Dim n_&
    
    Do
        
        n_ = Int(65536 * Rnd) 'ランダム性はあまりこだわらないでいい
        
    Loop Until ThisWorkbook.VBProject.VBComponents(namePrefix_ & n_) Is Nothing
    
    
    makeUniqueRndModuleName_ = namePrefix_ & n_
    
End Function






'変数名／関数名用に、一意になるような名前を生成する。

Private Function makeUniqueName_(prefix_$) As String
    
    Static nameSeed_&
    
    
    makeUniqueName_ = prefix_ & Replace(CDec(Now), ".", "x", , 1) & "x" & nameSeed_
    
    '{prefix_}{日付}x{時間}x{nameSeed_}となる。CDec()してるのは指数形式を回避するため。
    
    
    nameSeed_ = nameSeed_ + 1
    
End Function








'文字列コレクションを Join() する。

Private Function joinCollection_(c_ As Collection, delim_$) As String
    
    If c_.Count = 0 Then Exit Function
    
    Dim arr_()
    ReDim arr_(c_.Count - 1)
    
    Dim i&, value_
    For Each value_ In c_
        
        arr_(i) = value_
        
        i = i + 1
        
    Next
    
    joinCollection_ = Join(arr_, delim_)
    
End Function






Private Sub swap_(ByRef ref_a_, ByRef ref_b_)
    
    Dim tmp_: tmp_ = ref_a_
    ref_a_ = ref_b_
    ref_b_ = tmp_
    
End Sub



' ----------------------------------------------------










'動的コードホルダー関連 =============================================================================



'動的コード関連の整理を行う。

Private Sub cleanDynamicCode_()
    
    cleanTmpStdModule_
    
    cleanTmpClass_
    
    scanDynamicCodes_
    
End Sub





'メインコードホルダーからデリゲート辞書へ登録する --------------------------------------



'生成済み動的関数を走査して辞書に登録する。
'　キーとするコード文字列は、関数内の先頭１行目以降にコメントとして挿入されている。
'　改行がエスケープされているので、vbCrLf に戻す必要がある。

Private Sub scanDynamicCodes_()
    
    
    '動的コードホルダのインスタンスを生成する。
    
    Dim mainCodeHolder_ As Object
    Set mainCodeHolder_ = New DynamicCode 'instantiate_(cMainCodeHolderName)
    
    
    '動的コードホルダをスキャンし、動的メソッドを delegates_ に登録していく。
    
    Dim cm_ As CodeModule
    Set cm_ = getCodeModule_(cMainCodeHolderName, vbext_ct_ClassModule)
    
    Dim funcName_ '$
    For Each funcName_ In getFuncNames_(cm_, cFuncNameHeader & "*")
        
        
        registDelegates_ _
 _
            keys_:=commentsToKeyCodes_(cm_, CStr(funcName_)), _
 _
            action_:=Delegate.CNew(mainCodeHolder_, funcName_)
        
        
    Next
    
    
End Sub



Private Sub registDelegates_(keys_ As Collection, action_ As Delegate)
    
    Dim key_ '$
    For Each key_ In keys_
        
        Set delegates_(key_) = action_
        
    Next
    
End Sub



Private Function commentsToKeyCodes_(codeModule_ As CodeModule, funcName_$) As Collection
    
    Dim iFunc_&:    iFunc_ = codeModule_.ProcBodyLine(funcName_, vbext_pk_Proc) '関数のある行
    
    
    Dim keys_ As New Collection
    
    Dim i&: i = iFunc_ + 1
    Do
        
        Dim line_$:     line_ = codeModule_.Lines(i, 1)
        
        If Not line_ Like "'*" Then Exit Do
        
        
        Dim enclosed_$: enclosed_ = Mid(line_, 2, Len(line_) - 2)       '両サイドの「'」を抜く
         
        Dim key_$:      key_ = Replace(enclosed_, escChar_, vbCrLf)  '改行エスケープを戻す
        
        keys_.Add key_
        
        
        i = i + 1
        
    Loop
    
    
    Set commentsToKeyCodes_ = keys_
    
End Function



'-----------------------------------------







'動的コードの整理を行う --------------------------


'個別に一時作成した関数定義クラスの内容を、ホルダークラスに移動する。個別クラスは削除する。

'コメントとしてコード文字列を記述してあるが、
'　・一つのコード文字列は一行に収められる。改行は escapeCh_ にエスケープされている。
'　・複数行ある場合がある。（同じ内容の関数も、スペースの数などでコード文字列の表現が変わったりするため）
'　・必ずホワイトスペースの単一スペース化等を施した代表コードを１つ持つ。←実装後回し

Private Sub cleanTmpClass_() 'Optional tmpModules_ As Collection)
    
    
    'メインコードホルダーを取得（なければ作成）
    
    Dim cmDst_ As CodeModule
    Set cmDst_ = getCodeModule_(cMainCodeHolderName, vbext_ct_ClassModule)
    
    If Not cmDst_.find("Captures", 1, 1, -1, -1) Then
        
        cmDst_.AddFromString makeDynamicCodeHead_(cMainCodeHolderName)
        
    End If
    
    
    '個別コードモジュールを列挙
    
    Dim cmSrc_ As CodeModule
    For Each cmSrc_ In getCodeModules_(cTmpNameHeader & "*", vbext_ct_ClassModule) '関数名先頭を指定して検索する
        
        
        '個別コードモジュール（元）からホルダークラス（先）へ、動的生成されたメソッドをコピー
        
        Dim funcName_ '$
        For Each funcName_ In getFuncNames_(cmSrc_, cNameForDynamicCodeSearch)
            
            cmDst_.AddFromString cmSrc_.Lines( _
 _
                cmSrc_.ProcStartLine(CStr(funcName_), vbext_pk_Proc), _
 _
                cmSrc_.ProcCountLines(CStr(funcName_), vbext_pk_Proc) _
            )
            
        Next
        
        
        '個別コードモジュール（元）を削除
        
        ThisWorkbook.VBProject.VBComponents.Remove cmSrc_.Parent
        
        
    Next
    
    
End Sub


'------------------------------------------------------








'動的コードの宣言文字列を作成する。

Private Function makeDynamicCodeHead_(className_$) As String
    
    makeDynamicCodeHead_ = _
 _
        "private declare PtrSafe function VariantCopy Lib ""oleaut32"" ( dst_, src_ ) as long" & vbCrLf & _
        "private declare PtrSafe function VariantCopyInd Lib ""oleaut32"" ( dst_, src_ ) as long" & vbCrLf & _
 _
        "public Captures as variant" & vbCrLf & _
 _
        vbCrLf & _
 _
        "public function Instantiate() as object: set Instantiate = new {class}: end function" & vbCrLf
    
    
    makeDynamicCodeHead_ = Replace(makeDynamicCodeHead_, "{class}", className_, 1, 1)
    
End Function





'一時標準モジュールを削除する。

Private Sub cleanTmpStdModule_()
    
    On Error Resume Next
    
    ThisWorkbook.VBProject.VBComponents.Remove _
 _
        ThisWorkbook.VBProject.VBComponents(cTmpStdModuleName)
    
End Sub

















Public Sub teste(t_$)
    
    Dim c_ As New Collection
    
    Debug.Print scanSugarSyntaxes_(t_, c_)
    
    Enumerable(c_).ToAryAsEnumerable.ForEach "x => debug.print x"
    
End Sub







'糖衣構文のパース =======================================



'糖衣構文をパースする。

Private Function scanSugarSyntaxes_(srcTxt_$, additionalFuncs_ As Collection) As String
    
    If srcTxt_ = "" Then Exit Function
    
    scanSugarSyntaxes_ = srcTxt_
    
    
    'ブロックタグ左右とブロックの左・中・右の文を操作処理するデリゲートを登録する。
    
    Dim manipurations_ As New Collection
    
    manipurations_.Add "over root" '.add before:=1 を使用するためにダミーを入れとく
    
    
    'ブロック構文を登録された処理でパースする。
    
    man_push_ manipurations_, tagL_:="(", tagR_:=")", manipuration_:=Delegate.CNew(Me, "syntaxParse_EscapeClosing_")
    
    scanSugarSyntaxes_ = scanSyntax_Block_(scanSugarSyntaxes_, manipurations_, results_:=additionalFuncs_)
    
    man_push_ manipurations_, tagL_:="[", tagR_:="]", manipuration_:=Delegate.CNew(Me, "syntaxParse_AnonymousCollectionFunc_")
    man_push_ manipurations_, tagL_:="{", tagR_:="}", manipuration_:=Delegate.CNew(Me, "syntaxParse_AnonymousDictionaryFunc_")
    
    scanSugarSyntaxes_ = scanSyntax_Block_(scanSugarSyntaxes_, manipurations_, results_:=additionalFuncs_)
    
    
    ' @ として退避したカンマを元に戻す。くるしいな…。
    
    scanSugarSyntaxes_ = Replace(scanSugarSyntaxes_, "@", ",")
    
End Function






'ブロック構文でくるまれた内部をパースするアルゴリズム。
'　パース処理関数は manipuration_ に Delegate として渡せる。

Private Function scanSyntax_Block_(srcTxt_$, manipurations_ As Collection, results_ As Collection) As String
    
    If srcTxt_ = "" Then Exit Function
    
    scanSyntax_Block_ = srcTxt_
    
    Do While manipurations_.Count > 1
        
        scanSyntax_Block_ = scanSyntax_Block_inner_(scanSyntax_Block_, 1, manipurations_, results_)
        
        man_pop_ manipurations_
    Loop
    
End Function



Private Function scanSyntax_Block_inner_(srcTxt_$, ipos_&, manipurations_ As Collection, results_ As Collection) As String
    
    If srcTxt_ = "" Then Exit Function
    
    
    Dim tagL_$, tagR_$, manipuration_ As IFunc
    
    man_peek_ manipurations_, out_tagL_:=tagL_, out_tagR_:=tagR_, out_manipuration_:=manipuration_
    
    
    
    scanSyntax_Block_inner_ = srcTxt_
    
    
    
    
    'オブジェクトを囲うタグを検出し、再帰的にオブジェクト生成関数名と置換していく。
    
    Dim iL_&:   iL_ = InStr(ipos_, scanSyntax_Block_inner_, tagL_)
    
    If iL_ = 0 Then Exit Function
    
    
    scanSyntax_Block_inner_ = scanSyntax_Block_inner_(scanSyntax_Block_inner_, iL_ + 1, manipurations_, results_)
    
    
    Dim iR_&:   iR_ = InStr(iL_ + 1, scanSyntax_Block_inner_, tagR_)
    
    If iR_ = 0 Then Exit Function
    
    
    
    
    'ブロックを分割する
    
    Dim txtL_$: txtL_ = Left$(scanSyntax_Block_inner_, iL_ - 1)
    
    Dim txtM_$: txtM_ = Mid$(scanSyntax_Block_inner_, iL_ + 1, iR_ - (iL_ + 1))
    
    Dim txtR_$: txtR_ = Mid$(scanSyntax_Block_inner_, iR_ + 1)
    
    
    
    
    'ブロック内部のテキストに対し、操作処理関数スタックに存在する manipuration_ を処理する。もっとシンプルにできないかね…。
    
    man_pop_ manipurations_
    
    If manipurations_.Count > 1 Then txtM_ = scanSyntax_Block_inner_(txtM_, 1, manipurations_, results_)
    
    man_push_ manipurations_, tagL_, tagR_, manipuration_
    
    
    
    'ブロックの左・中・右の文を操作処理し、結果として返す。
    
    scanSyntax_Block_inner_ = manipuration_.xExec04(txtL_, txtM_, txtR_, results_)
    
    
End Function



'操作処理関数スタック -----------------------

Private Sub man_peek_(manipurations_ As Collection, ByRef out_tagL_$, ByRef out_tagR_$, ByRef out_manipuration_ As IFunc)
    
    out_tagL_ = manipurations_(1)
    out_tagR_ = manipurations_(2)
    Set out_manipuration_ = manipurations_(3)
    
End Sub

Private Sub man_pop_(manipurations_ As Collection)
    
    manipurations_.Remove 1
    manipurations_.Remove 1
    manipurations_.Remove 1
    
End Sub

Private Sub man_push_(manipurations_ As Collection, tagL_$, tagR_$, manipuration_ As IFunc)
    
    manipurations_.Add manipuration_, before:=1
    manipurations_.Add tagR_, before:=1
    manipurations_.Add tagL_, before:=1
    
End Sub

'--------------------------------------------




'括弧のパース ---------------------------------------------------------

Public Function syntaxParse_EscapeClosing_(txtL_, txtM_, txtR_, results_) As Variant
    
'    syntaxParse_EscapeClosing_ = txtL_ & "(" & txtM_ & ")" & txtR_
    syntaxParse_EscapeClosing_ = txtL_ & "(" & Replace(txtM_, ",", "@") & ")" & txtR_
    
End Function

'----------------------------------------------------------------------



'匿名オブジェクトのパース ---------------------------------------------

'匿名オブジェクトをパースする。
'　Collection と Dictionary を生成する構文が存在する。
'　内部で、オブジェクトを生成する関数を動的に生成する仕組み。
'　[ x = 1, y = "str" ] などが、Collection 生成関数名に置き換わる。
'　{ x = 1, y = "str" } などが、Dictionary 生成関数名に置き換わる。
'　{ x, y } は { x = x, y = y } と同じ。
'　ネスト可能。正しく閉じられていない場合は未定義なので注意（大概の場合はまったくパースされないと思う）。
'　変数名のチェックなどもしていないので、書く側で注意して使用する。

' objType_
'　1 : collection   [ , ]
'　2 : dictionary   { , }
'　※ Array() も考えられるが、それはそのまま書いてもらえばいい

'匿名オブジェクトは、キー「"ay"」／バリュー「 Enumerable.AnonymousIdentify への参照」のメンバを持ち、通常の Collection/Dictionary と区別する。

'タグ文字／パラメータ区切り文字は、ＶＢＡ構文で使用されない文字でなければならない。
'　複雑な検索はしていないため。文脈で意味を考える等はしていない。


Public Function syntaxParse_AnonymousCollectionFunc_(txtL_, txtM_, txtR_, additionalFuncs_) As Variant
    
    
    '関数名をランダムに生成する。
    
    Dim funcName_$: funcName_ = makeUniqueName_(cCollNameHeader)
    
    
    '匿名オブジェクト生成関数文字列を構築
    
    Dim funcs_ As Collection
    Set funcs_ = additionalFuncs_
    
    funcs_.Add makeInstantiateFunctionString_ _
 _
        (ref_funcName_:=funcName_, propertyList_:=CStr(txtM_), typeName_:="Collection", separator_:=",", needSwap_:=True)
    
    
    '関数名に置換後のコード文字列
    
    syntaxParse_AnonymousCollectionFunc_ = txtL_ & funcName_ & txtR_
    
End Function

Public Function syntaxParse_AnonymousDictionaryFunc_(txtL_, txtM_, txtR_, additionalFuncs_) As Variant
    
    
    '関数名をランダムに生成する。
    
    Dim funcName_$: funcName_ = makeUniqueName_(cDictNameHeader)
    
    
    '匿名オブジェクト生成関数文字列を構築
    
    Dim funcs_ As Collection
    Set funcs_ = additionalFuncs_
    
    funcs_.Add makeInstantiateFunctionString_ _
 _
        (ref_funcName_:=funcName_, propertyList_:=CStr(txtM_), typeName_:="Dictionary", separator_:=",", needSwap_:=False)
    
    
    '関数名に置換後のコード文字列
    
    syntaxParse_AnonymousDictionaryFunc_ = txtL_ & funcName_ & txtR_
    
End Function


'匿名オブジェクト生成関数を文字列で作成する。

Private Function makeInstantiateFunctionString_(ByRef ref_funcName_$, propertyList_$, typeName_$, separator_$, needSwap_ As Boolean) As String
    
    Dim props_$(): props_ = Split("ay=grouping.anonymousidentify" & separator_ & propertyList_, separator_)
    
    
    Dim lines_  As New Collection   '関数内の文
    Dim params_ As New Collection   '引数リスト
    Dim values_ As New Collection   '値リスト
    
    Dim prop_ '$
    For Each prop_ In props_
        
        prop_ = Trim$(prop_)
        
        If prop_ <> "" Then
            
            Dim nameAndValue_$(): nameAndValue_ = Split(prop_, "=", 2)
            
            Dim key_$:      key_ = Trim(nameAndValue_(0))
            Dim value_$:    value_ = nameAndValue_(UBound(nameAndValue_))
            
            
            params_.Add key_
            values_.Add value_
            
            
            Dim p0_$, p1_$
            
            p0_ = """" & key_ & """"
            p1_ = key_
            
            If needSwap_ Then swap_ p0_, p1_
            
            lines_.Add ref_funcName_ & ".Add " & p0_ & "," & p1_
            
        End If
        
    Next
    
    
    Dim funcNameWithParams_$
    funcNameWithParams_ = ref_funcName_ & "(" & joinCollection_(params_, ",") & ")" ' dict000x000x0( x, y ) のような引数リスト付きの名前にする
    
    Dim funcNameWithValues_$
    funcNameWithValues_ = ref_funcName_ & "(" & joinCollection_(values_, "@") & ")" ' dict000x000x0( 1, x + 2 ) のような代入リスト付きの名前にする
    
    
    
    Dim strFunc_$: strFunc_ = _
 _
        "private function {funcName(a,b,c)} as {type}:" & _
 _
            "set {funcName} = new {type}:" & _
 _
            "{line:line:line}:" & _
 _
        "end function"
    
    strFunc_ = Replace(strFunc_, "{funcName(a,b,c)}", funcNameWithParams_, 1, 1)
    strFunc_ = Replace(strFunc_, "{funcName}", ref_funcName_, 1, 1)
    strFunc_ = Replace(strFunc_, "{type}", typeName_, 1, 2)
    strFunc_ = Replace(strFunc_, "{line:line:line}", joinCollection_(lines_, ":"), 1, 1)
    
    
    
    ref_funcName_ = funcNameWithValues_
    
    makeInstantiateFunctionString_ = strFunc_
    
End Function


'---------------------------------------------








' a ? b : c 構文のパース ---------------------

Private Function scanShortBranch_(srcTxt_$, factoryFuncs_ As Collection) As String
    
    scanShortBranch_ = srcTxt_
    
    
    
    
    
End Function

'---------------------------------------------
'
'
'Private Sub syntaxShortIfBranchFunc__CallFuncer03(txtL_, txtM_, txtR_, this_ As IFuncer, result_)
'
'    Dim factoryFuncs_ As Collection:    Set factoryFuncs_ = this_.Object(2)
'
'
'    txtL_ = Left(txtL_, Len(txtL_))
'    txtR_ = Right(txtR_, Len(txtR_))
'
'
'    Dim ref_funcName_$: ref_funcName_ = makeUniqueName_("if")                   '関数名をランダムに生成する。
'
'    factoryFuncs_.Add makeShortIfFunctionString_(ref_funcName_, CStr(txtM_))    '匿名オブジェクト生成関数の文字列
'
'    result_ = txtL_ & ref_funcName_ & txtR_                                     '関数名に置換後のコード文字列
'
'End Sub
'
'Private Function makeShortIfFunctionString_(ByRef ref_funcName_$, propertyList_$) As String
'
'
'    Dim lines_ As New Collection    '関数内の文
'    Dim params_ As New Collection   '引数リスト
'    Dim values_ As New Collection   '値リスト
'
'
'    params_.Add "a"
'    values_.Add "b"
'
'
'    Dim funcNameWithParams_$:   funcNameWithParams_ = ref_funcName_ & "(" & joinCollection_(params_, ",") & ")" ' ay000x000x0( x, y ) のような引数リスト付きの名前にする
'
'    Dim funcNameWithValues_$:   funcNameWithValues_ = ref_funcName_ & "(" & joinCollection_(values_, ",") & ")" ' ay000x000x0( 1, x + 2 ) のような代入リスト付きの名前にする
'
'
'    makeShortIfFunctionString_ = _
' _
'        "private function " & funcNameWithParams_ & "：" & _
' _
'            "if a then b else c" & "：" & _
' _
'        "end function"
'
'
'    ref_funcName_ = funcNameWithValues_
'
'
'End Function
'
'











'初期化・終期化 =====================================================


Private Sub Class_Initialize()
    
    initVariable_
    
    cleanDynamicCode_
    
End Sub

Private Sub Class_Terminate()
    
    cleanDynamicCode_
    
End Sub





