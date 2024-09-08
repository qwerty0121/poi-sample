# poi-sample

## 概要

Apache POI を利用した処理のサンプル実装。

## 実行手順

以下のコマンドを実行すると、 ".output" ディレクトリに Excel ファイルが出力される。

```bash
mvn clean package

mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.SheetCopySample"
```
