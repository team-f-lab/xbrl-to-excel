# XBRL to Excel

ExcelにXBRLの値を読み込む処理のデモです。    
事前に決算サマリExcelテンプレートを作成しておき、決算書が発表されたらテンプレートにXBRLの値を読み込む・・・といった用途に使用できます。

## Dependencies

* [`coarij`](https://github.com/chakki-works/CoARiJ)
  * EDINET上で公開されている決算書の情報を取得するのに使用します。
* [`xbrr`](https://github.com/chakki-works/xbrr)
  * XBRLを読み取るのに使用します。
* [`openpyxl`](https://bitbucket.org/openpyxl/openpyxl/src)
  * Excelを読み込む/値を書き込むのに使用します。

## Execution

依存ライブラリのインストール

```
pip install -r requrements.txt
```

XBRL=>Excelの実行

```
python execute_convert.py
```
