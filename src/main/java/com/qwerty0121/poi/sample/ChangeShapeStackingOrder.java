package com.qwerty0121.poi.sample;

import java.io.IOException;
import java.util.Collections;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;

import com.qwerty0121.poi.utils.PoiSampleUtils;

/**
 * 図形の重なり順を変更するサンプル
 */
public class ChangeShapeStackingOrder {

  public static void main(String[] args) throws IOException {
    try (var workbook = PoiSampleUtils.createWorkbook("図形重なり順変更テンプレート.xlsx");) {
      // "テスト"シートを取得
      var sheet = workbook.getSheet("テスト");

      // 図形の重なり順を逆にする
      reverseShapeStackingOrder(sheet);

      PoiSampleUtils.writeWorkbook(workbook, "図形重なり順変更.xlsx");
    }
  }

  /**
   * 図形の重なり順を逆にする
   * 
   * @param sheet シート
   */
  private static void reverseShapeStackingOrder(Sheet sheet) {
    // --- 注意事項 ---
    // 今回のサンプルではテンプレートのExcelファイルにCTTwoCellAnchorのみ含まれていることを前提としている。
    // そのため、CTOneCellAnchorやCTAbsoluteAnchorが含まれている場合は、別途処理を追加する必要がある。
    // CTDrawingからCTTwoCellAnchorを追加/削除することで図形の重なり順を変更できる。
    // それ以外の方法では出力時にエラーとなる、出力したファイルが破損する、などの問題が発生する。
    // また、CTDrawingから削除したCTTwoCellAnchorインスタンスを再度CTDrawingに追加するとXmlValueDisconnectedExceptionが発生する。
    // そのため、一度XML文字列に変換し、そのXML文字列から再度CTTwoCellAnchorインスタンスを生成している。

    var drawing = sheet.getDrawingPatriarch();
    if (!(drawing instanceof XSSFDrawing xssfDrawing)) {
      throw new RuntimeException("シートからのDrawingの取得に失敗しました。");
    }

    // CTDrawingを取得
    var ctDrawing = xssfDrawing.getCTDrawing();

    // CTTwoCellAnchorリストを取得
    var twoCellAnchorList = ctDrawing.getTwoCellAnchorList();

    // CTTwoCellAnchorリストをコピーし、変更後の順序に変更する
    // ※今回は逆順にする
    var newOrderTwoCellAnchorList = twoCellAnchorList.stream().map(twoCellAnchor -> {
      // 後の処理でCTTwoCellAnchorを削除するため、CTTwoCellAnchorインスタンスをコピーする
      // NOTE: これをしないとCTDrawingに追加する際にXmlValueDisconnectedExceptionが発生してしまう

      // XML文字列に変換
      var xmlText = twoCellAnchor.xmlText();
      // 再度CTTwoCellAnchorに変換
      try {
        return CTTwoCellAnchor.Factory.parse(xmlText);
      } catch (XmlException e) {
        throw new RuntimeException(e);
      }
    }).collect(Collectors.toList());
    Collections.reverse(newOrderTwoCellAnchorList);

    // 一旦、全ての図形を削除する
    twoCellAnchorList.clear();

    // 逆順にした図形を追加する
    newOrderTwoCellAnchorList.forEach(twoCellAnchor -> {
      ctDrawing.addNewTwoCellAnchor().set(twoCellAnchor);
    });
  }
}
