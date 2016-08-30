# DbTbl

このツールは、ExcelにDBの内容を取得して来たり、そのDBを更新することができる
ちょっと危ないツールです。

# Build

VBAのimportに、igeta（[@igeta](http://twitter.com/igeta)）さんのAriawase(vbac.wsf)を使用させていただいております。
（このツールが無ければgitHubに公開しようとは思わなかった。）多謝。

下記Ariawaseのライセンス

https://github.com/vbaidiot/Ariawase#license

以下でビルドします。

> build.bat

# Requirement

必須：

* Windows
* Excel（xlsmなのでExcel 2007以降の最近のやつ）
* ODBCプロバイダの設定
  * ※Excelのbit版（32bit or 64bit）バージョンと同等のもの（の設定）が必要なハズです。

現状postgresのODBCみ動作確認しています。

# How to use

## 設定シート

### カラム

|位置|内容|設定例|
|------|------|---------|
|D1   |接続文字列|DSN=PostgreSQL_local|

### ボタン

|キャプション|内容|バインド関数名|備考|
|------------------|------|---------------------|------|
|#×のシートの削除|A列が×の行を削除＋存在するシートの削除|DeleteSheets||
|件数取得|件数を取得する|GetAllCount||
||GetAllScript||
|テーブル一覧取得|テーブルの一覧をシートに出力します|GetAllTablesList||
|テーブル定義取得|テーブル定義シートを作成します|GetAllTables|てんぷれシートからコピー後定義反映|
|全取得||GetAllData|全スクリプト作成|設定のシートにある定義のすべての「Excel→script」ボタン動作|

シート名＝テーブル名の予定でしたがExcelの制限でテーブル名が31文字を超えるものは、少し切れてます。

## てんぷれ

※このシートがテンプレートになり各シートが生成される

|位置|内容|設定例|備考|
|------|------|---------|------|
|D1|接続文字列|DSN=PostgreSQL_local|設定シートよりコピーされる|
|D2|テーブル名|table_name||
|D3～xx3|項番|1||
|D4～xx4|項目名称|カラム名|現状フィールド名と同じ|
|D5～xx5|フィールド名|field_name||
|D6～xx6|タイプ|integer||
|D7～xx7|桁||未使用|
|D8～xx8|バイト||未使用|
|D9～xx9|PK|○|プライマリキーの場合「○」|
|D10～xx10|Not Null|○|Not Nullの場合「○」|
|D11～xx11|Order||未使用|
|D12～xx12|FK||未使用|
|D13～xx13|条件||未使用|

### ボタン

|キャプション|内容|バインド関数名|備考|
|---------|----|-----------|----|
|Clear Excel|21行目以降を削除|ExcelClear||
|Excel→script|シートの内容をsqlとcsvでファイルに出力|CreateBat||
|DB→Excel|DBからデータ取得|Export|21行目以降に出力|
|Excel→DB|DBを更新|UpdateDB|21行目以降で#が続く限り、データ更新を試みる|

# Demonstration

postgresが入っていない場合はインストールしてください。

> createdb testdb

testdb（名前を合わせれば何でもOK）を作成します。

postgresとODBCドライバをインストールして下さい。

下記からサンプルのpostgresデータを取得して下さい。

http://www.postgresqltutorial.com/postgresql-sample-database/

dvdrental.tar

> tar xvf dvdrental.tar

解凍後、

restore.sqlを編集してください。

```
$$PATH$$ →　適切なパスへ
```

> psql -l 127.0.0.1 -d testdb -U postgres -f restore.sql

restore.sqlをpsqlでデータを流してください。

下記のようにしてODBCドライバを組み込んでください。

設定するODBCはbitバージョンをExcelのbitバージョンに合わせてください

設定例

32bit windows、64bit windows　の各bit版ドライバ

> odbcconf.exe /A {CONFIGDSN "PostgreSQL Unicode" "DSN=PostgreSQL_local|DATABASE=testdb|SERVER=127.0.0.1|PORT=5432|UID=postgres|PWD=postres"}

64bit windows の　32bitドライバ

> %systemdrive%\Windows\SysWoW64\odbcconf.exe /A {CONFIGDSN "PostgreSQL Unicode" "DSN=PostgreSQL_local|DATABASE=testdb|SERVER=127.0.0.1|PORT=5432|UID=postgres|PWD=postres"}

動作のデモ

[![DBTblエクセルデモ](https://i9.ytimg.com/vi/Q9I2nwsZ-yM/1.jpg?sqp=CNSjl74F&rs=AOn4CLCyimYBieiR8AKgMT2YzEPZqOh_aQ&time=1472582134487)](https://youtu.be/Q9I2nwsZ-yM?cc_load_policy=1&loop=1&v=qL9GBq0V6Q4:embed:cite)

# History

* 2016/08/30 Initial Commit

# Known isssue

* 日付型がDBに書き出せない
