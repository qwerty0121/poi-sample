# poi-sample

## 概要

Apache POI を利用した処理のサンプル実装。

## 実行手順

以下のコマンドを実行すると、 ".output" ディレクトリに Excel ファイルが出力される。

```bash
mvn clean package

# シートコピー
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.SheetCopySample"

# シートコピー(別ワークブック)
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.SheetCopyToOtherWorkbookSample"

# 図形テキストの設定
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.ShapeTextSettingSample"

# 画像追加
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.AddImageSample"

# 図形非表示
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.HideShapeSample"

# 図形削除
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.RemoveShapeSample"

# 図形重なり順変更
mvn exec:java -Dexec.mainClass="com.qwerty0121.poi.sample.ChangeShapeStackingOrder"
```
