# ProBoarder Excel とは

ProBoarder Excel は、マウスが必要なExcel操作をすべてキー入力によるコマンドで実行できるExcelアドインです。

2020/10/14現在、主なコマンドは以下のとおりです。

|機能|コマンド名|補足事項|
|-|-|-|
|シート追加|add sheet|-|
|アクティブシート名を変更|change sheet name|-|
|アクティブシートを削除|delete sheet|-|
|行幅を自動的に調整する|autofit row|-|
|列幅を自動的に調整する|autofit col|-|
|行幅と列幅を自動的に調整する|autofit|-|
|セルA1を選択する|select a1|-|
|全シートのセルA1を選択する|select a1 all sheet|-|
|ヘルプページを表示する|help|-|
|CSVファイルを選択してアクティブセルに貼り付ける|paste csv from file|-|
|タブ区切りファイルを選択してアクティブセルに貼り付ける|paste tsv from file|-|
|CSVファイルを選択して文字列としてアクティブセルに貼り付ける|paste csv from file as string|-|
|タブ区切りファイルを選択して文字列としてアクティブセルに貼り付ける|paste tsv from file as string|-|
|拡大率を変更する|zoom|-|
|選択範囲に合わせて拡大縮小する|zoom fit selection|-|
|右方向（列）にオートフィル|autofill col|-|
|下方向（行）にオートフィル|autofill row|-|
|並び替えダイアログを表示する|show dialog sort|-|
|シートの移動またはコピーダイアログを表示する|show dialog workbook move|-|
|選択範囲のオートフィルタをオンにする|set autofilter on|-|
|選択範囲のオートフィルタをオフにする|set autofilter off|-|
|選択範囲に行を挿入する|insert row|-|
|選択範囲に列を挿入する|insert col|-|
|選択範囲の行を削除する|delete row|-|
|選択範囲の列を削除する|delete col|-|
|アクティブワークブックを読み取り専用に変更する|change file access readonly|-|
|アクティブワークブックを書き込み可能に変更する|change file acces readwrite|-|
|Excelアプリケーションを終了する|application quit|-|
|アドイン管理ダイアログを表示する|show dialog addin manager|-|
|選択範囲のセル幅を変更する|set cell width|-|
|選択範囲のセル高さを変更する|set cell height|-|
|アクティブシート全体のセルの幅を変更する|set entire cell width|-|
|アクティブシート全体のセルの高さを変更する|set entire cell height|-|
|選択範囲のセルの背景色を変更する|set cell color|-|
|選択範囲のセルの文字色を変更する|set font color|-|
|選択セルをリンクされた図として貼り付ける|paste as linked picture|-|


## インストール方法

1. ProBoarderExcel.xlam をダウンロードし、ローカルディスクに保存します。

2. Excelを起動し、アドイン管理ダイアログを表示します。

![](images/ExcelMenuDevelopment.png)

![](images/2020-10-12-17-10-49.png)

3. 参照をクリックし、ProBoarderExcel.xlam を選択します。

![](images/2020-10-12-17-11-52.png)

4. OKボタンをクリックします。

6. アドイン管理ダイアログに「Proboarderexcel」にチェックが入っていることを確認し、OKボタンをクリックします。

![](images/2020-10-12-17-12-09.png)


7. ワークシート画面で「`Ctrl + D`」 を押して、ProBoarderExcelウィンドウを表示します。

![](images/ExcelMenuDevelopment.png)

![](images/2020-10-12-17-12-33.png)


8. コマンドを入力して使用します。


## ProBoarder Excel の使い方

### 1. ProBoarder Excel ウィンドウ表示方法

[ Ctrl + D ] を 押して、ウィンドウを表示します。

![](images/2020-10-12-17-12-33.png)

### 2. コマンド入力

テキストボックスにコマンドを入力し、Enterを押すと実行します。

定義されていないコマンドが入力された場合、何も動作しません。




### 3. コマンド一覧表示

テキストボックスが選択された状態で、[ Alt + ↓ ] を入力すると、一覧が表示されます。

