
# vbaEnumeration

## これは何

ＶＢＡのクラス及び標準モジュールにより、ＬＩＮＱに近い機能を実現するものです。  
以下の機能が実装されています。   

* Delegate クラスによるクラスオブジェクトメソッドのデリゲート化   
* FuncPointer クラスによる標準モジュール関数の関数オブジェクト化   
* Capt クラスによる疑似ラムダ式（処理内容は文字列で記述）のデリゲート化   
* Enumerator クラスによるオブジェクトの For Each 列挙   
* Enumerator クラスメソッドからのＬＩＮＱ的オペレータの使用（高階関数的に疑似ラムダ式・デリゲート・関数オブジェクトを渡すことが可能）   

さらにおまけとして   

* Ary クラスによる配列のオブジェクト化（値渡しではなく参照渡しを可能とする）   

## どうやって使う

各種モジュールをＶＢＥにインポートしてください。  
参照設定は通常の Excel/Access の設定に加え、以下を必要とします。  
* Microsoft Scripting Runtime  

簡単な使用例

		dim f as IFunc                                    'デリゲート／関数オブジェクトのインターフェース    
		set f = Delegate.CNew( new Class1, "MethodA" )    ' Delegate はデフォルトインスタンス   
		Debug.Print f( x, y )                             ' f.Exec( x, y ) の省略形   
		   
		set f = FuncPointer.CNew( AddressOf FuncA )       ' Sub は未対応   
		Debug.Print f( x, y )   
		  
		set f = Capt( " x, y =>> x + y " )                '疑似ラムダ式   
		Debug.Print f( 1, 2 )                             ' 3 と表示される   
		   
		Dim i&: i = 3
		set f = Capt( " x : c =>> x * c ", i )            ' c は i を疑似的にキャプチャし、3 となる
		Debug.Print f( 10 )                               ' 30 と表示される   
		   
		Dim c as new Collection   
		c.add 1   
		c.add 2   
		c.add 3   
		
		Dim i   
		For Each i in Enumerable(c).qWhere( " x =>> x > 1 " ).qSelect( " x =>> x + 10 " )   
		    Debug.Print i;   
		Next   
		' 12 13 が表示される
		
		Dim arr as Ary
		set arr = Ary.CNew.Alloc( 10, vbLong )
		arr(0) = 1&
		arr(9) = 10&
		arr.DimDef(10).DimDefBound(1,5).DimAlloc( vbString )
		arr(0,1) = "first element"
		arr(9,5) = "last element"

## 注意点

まだまだ未完成／未整理です。   
エラー処理はほぼ入れてないです。テストもしていないに等しいです。   
デリゲートや関数オブジェクトは、引数の数や型（現状は Variant のみ対応）を間違えて指定するなどで容易にエクセルが落ちます（！！）。   
GitHub の使い方がよくわかりません。勉強していきます。   
コードが読みづらいと言われます。今日もさんざん考えさせられました。    

## ＬＩＮＱ的オペレータ一覧

selector 等には、文字列のコード式／IFunc を実装するもの／標準モジュール関数のアドレスを渡すことができます。   
また Optional となっているものを省略すると、大抵は流れてくる値を素通しします。   
* qSelect( selector )
* qWhere( predicate )
* qSelectMany( Optional collectionSelector, Optional resultSelector )
* qGroupBy( Optional keySelector, Optional elementSelector, Optional resultSelector )
* qSkip( count as long )
* qTake( count as long )
* qSpan( count as long )
* Sum
* Count
* ForEach( expression )
* ToAry( Optional baseIndex as long )
* ToCollection
* ToDictionay( Optional keySelector, Optional elementSelector )
* ToLookUp( Optional keySelector, Optional elementSelector )
* その他順次追加

## 疑似ラムダ式について

* 疑似ラムダ式は、動的にクラスモジュールを生成する手法により実現しています。   
ＶＢＥから実行すると、あわただしくウィンドウが生成されていく過程が見て取れます。   
ただし、一度生成すれば ZTmpxxxx クラス（一時コードストック）を経由して DynamicCode クラス（永続コードストック）にストックされ、次からはそのコードが再利用されるようになっています。   
そうなれば、デバッグトレースも可能です（自動生成される名前がわかりにくいことこの上ない感じになりますが）。   
また ZTmpxxxx や DynamicCode クラスモジュールは一時クラス扱いなので、削除（破棄）しても再生成されます。   
* 「;」は改行扱いされます。
* &gl;result&gt; = x と書くと戻り値になります。ただし、一行のみの疑似ラムダ式に関しては、行頭に「>」を入れることで式の結果が戻り値になります。   
疑似ラムダ式は x => x が通常の構文になりますが、x =&gt;&gt; x とすることで「=&gt;&gt;」は「=&gt;」＋ 行頭の「&gt;」と解釈されます。
* 引数を必要としない場合は、「x,y =&gt;」の部分ごと省略できます。つまり、Capt("1") も正当な構文と解釈されます。
* 文字列は「' '」でくくります。$'text{value}text' とすれば変数を埋め込むことができます。
* 疑似匿名オブジェクトが使用できます。{ x=1, y } 等で匿名 Dictionary、[ x='a', set y=[a,b] ] 等で匿名 Collection が生成されます。
* 匿名オブジェクトを qGroupBy/ToLookUp のキーとして使用した場合、ＬＩＮＱのようにメンバ内容を元にグループ化されます。
グループ化は、匿名オブジェクトを文字列化して Dictionary のキーにすることで行っています（匿名入れ子にも対応）。   
そのため ToLookUp のキーを匿名オブジェクトとした場合には、辞書引きする際に Grouping.ToAyKey(key) 関数を通して文字列化しておく必要があります。
* 現状では、引数／戻り値の型はすべて Variant 型としなければなりません。

## どうやっているのか

* Delegate/FuncPointer は DispCallFunc() を使用しています。   
そのため、かなりデリケートです（デリゲートだけに！）。引数の数や型をしくじったたけでハングアップします。何とかしたい。
* 列挙は IEnumVariant を CoTaskMemAlloc() で生成して渡しています。メンバメソッドを xCom 標準モジュールで定義し、v-table に登録してＣＯＭオブジェクトとしてふるまわせています。のですが、できたらクラスモジュールのメソッドを登録したかった。ＶＢＡクラスメソッドの引数／戻り値は素直に渡されないので、IEnumVariant のメンバ関数としてコールさせられないようです。一つだけ標準モジュールが存在している理由はここにあります。
* その他、デフォルトプロパティとデフォルトインスタンスを使用しています。これらはＶＢＥエディタからは指定できないようで、エクスポートしたファイルを書き換えて指定しています。

## 列挙可能なオブジェクト

* IEnumVARIANT を返せるオブジェクト（ _NewEnum メソッド／プロパティを持つオブジェクト）
* Ary オブジェクト（素の Array は現在対象外）←ただし、列挙中にデバッグトレースをかけると特定の場所でエクセルが落ちる…
* 現在未実装ですが、レコードセットなども対応予定

## モジュール一覧

* IFunc
* Delegate　　　　　　　 IFunc インターフェースを実装
* FuncPointer　　　　　　IFunc インターフェースを実装　名前は FuncObject とかのほうがいいだろうか
* Capt　　　　　　　　　 肥大化ぎみ、RegExp の使用を再検討中…orz
* Enumerable　　　　　　.From() しかメンバがないので、要らないかも？
* Enumerator
* EnumOperatorProcs　　オペレーターの処理が記述されている
* Grouping　　　　　　　Collection で代用してもいいかも
* Ary　　　　　　　　　　インターフェースに難ありかと
* xCom　　　　　　　　　ＣＯＭ関連のユーティリティと EnumVariant の実体実装　（標準モジュール）

