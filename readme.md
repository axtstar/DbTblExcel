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
* Excel(xlsmなのでExcel 2007以降の最近のやつ)
* ODBCプロバイダの設定
  * ※Excelのbit（32bit/64bit）バージョンと同等のもの（の設定）が必要なハズです。

現状postgresのODBCみ動作確認しています。

# How to use

設定シートの
　D1が接続文字列の記述です。
　　記述例：DSN=PostgreSQL_staging

　E1にあるテーブル定義取得を押すと、接続したデータベースのスキーマを見に行って、テーブルを取得してシートをいっぱい追加します。
　シート名＝テーブル名の予定でしたがExcelの制限でテーブル名が31文字を超えるものは、少し切れてます。
　
　C1の件数取得で先ほど取得したテーブル一覧に件数を取得します。
　
　各シートでDBからのデータ取得とDBへのデータ更新ができます。

# History

* 2016/06/27 Initial Commit

# Known isssue

* 日付型がDBに取り込めない