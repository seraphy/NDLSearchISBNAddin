# これは何か？

ISBNをキーに国会図書館が公開する図書の書誌情報を抽出する、VBAマクロによるアドインです。

# 使い方

アドイン「ISBN検索.xlam」を実行すると、ExcelのリボンバーにISBN検索という項目が増えます。

このメニュー項目を選択すると、

- ISBN検索
- フォーム設定
- About

の３つのアイコンが表示されます。

## 設定方法

フォーム設定で検索開始する行番号、ISBNの列番号等を指定します。


また、ここで接続方法としてWinINetかWinHTTPのどちらかを選択できます。


なお、設定項目のうち、シートの設定はワークブックの「カスタムプロパティ」と「レジストリ」の双方に、接続の設定はレジストリに保存されます。

(※ シート設定をレジストリに保存するのは次回新規ワークブックを開いたときに前回の設定値を反映させるため)

レジストリは「HKEY_CURRENT_USER\Software\VB and VBA Program Settings\SearchISBNAddin」下に記録されます。

これらの項目はアドイン側から消すことはないので不要であれば手動で消してください。

※ Proxyのパスワードも平文で書き込まれるのでレジストリの扱いには注意してください。


## 実行

「ISBN検索」ボタンを押下すると、フォーム設定に従って現在のアクティブシートのISBN列を読み取り、
その検索結果を書き込んでゆきます。

検索できなかった場合はISBNセルの背景色を青に変更して検索できなかったことを示します。

ISBNとして妥当でない番号が入力されている場合はISBNセルの背景色を赤に変更します。
(この場合、検索は行われません。)

すでにタイトル列がある行はスキップされます

指定された行番号から、使用されているセルの全行までの、すべてのISBN列を走査します。


## 終了

Excelを終了すれば、このアドインは終了しリボンのアイコンも消えます。

Excelのアドインフォルダに入れて常時起動に設定することもできます。


# 注意事項

国立国会図書館サーチのWebAPIは非営利での利用を前提としており、営利目的(無償提供/有償提供に関わらず)での使用では申請が必要です。

(非営利の個人利用でも継続的に使う場合には利用状況を把握するために申請が望ましいとのことです。)

詳しくは、国立国会図書館サーチのAPIのホームページを確認ください。

http://iss.ndl.go.jp/

また、本アドインが使用しているAPIは、以下のOpenSearch形式のものです。

> http://iss.ndl.go.jp/api/opensearch?isbn=


## 元ネタ

このアドインは、LOD Challenge 2013 書誌検索シート(原田隆史 同志社大学)のマクロをベースにしています。

http://lod.sfc.keio.ac.jp/challenge2013/show_status.php?id=a072

※ ただし、オリジナルとは使用しているAPIは異なります。


## ライセンス

MIT


## その他

国立国会図書館サーチは必ずしも市販されている全ての本のISBNを検索できるわけではないようです。

私が試したところでは9割がたは検索できているので、もし検索できない書籍があればAmazon等で手動で検索すれば、だいたい埋められます。

----

# ビルド方法

エクセルのアドインなので本来バイナリですが、ソース管理のために、Ariawaseのvbacツールを使ってソースを展開しています。

https://github.com/vbaidiot/Ariawase


```
    cscript vbac.wsf decombine
```

でbinフォルダ上にある*.xlamがソースとしてsrcフォルダに展開されます。


しかし、このツールからはxlamは直接再構築できないので、一旦、xlsmとして再構築させてから、xlamに変換(アドインとして保存しなおし)する必要があります。

(フォルダ名を*.xlsmに変えてcombineします。)

参照設定として、"**Microsoft XML, v6**"の設定も必要です。

また、生成されたxlsmにはリボン設定等がないので、

リボンのための
- _rels/.rels
- customUI/customUI.xml

の追加が必要です。

(このXMLにリボンのアイコン名や表示名、ハンドラの名前などが定義されています。)

これは*.xlamを一旦zip展開後に、上記ファイルを上書きして、再度zipで固め直すことで、アドインを呼び出すリボンの設定が付与できます。

また、Excelのアドインダイアログで表示されるアドイン名は、エクスプローラからxlamファイルを選択してコンテキストメニューを開き、プロパティから詳細タブを選択し、「タイトル」プロパティを設定することで変更できます。
