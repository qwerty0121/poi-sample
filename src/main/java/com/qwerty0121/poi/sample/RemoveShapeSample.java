package com.qwerty0121.poi.sample;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import com.qwerty0121.poi.utils.PoiSampleUtils;

/**
 * 図形を削除するサンプル
 */
public class RemoveShapeSample {

  public static void main(String[] args) throws IOException {
    try (var workbook = PoiSampleUtils.loadTemplateWorkbook("図形削除サンプルテンプレート.xlsx");) {
      // "テスト"シートを取得
      var sheet = workbook.getSheet("テスト");

      // 図形を削除する
      removeShape(sheet, "shape");

      // 図形グループを削除する
      removeShape(sheet, "shape-group");

      // 画像を削除する
      removeShape(sheet, "picture");

      PoiSampleUtils.writeWorkbook(workbook, "図形削除サンプル.xlsx");
    }
  }

  /**
   * 図形を削除する
   * 
   * @param sheet     シート
   * @param shapeName 削除対象の図形の名前
   */
  private static void removeShape(Sheet sheet, String shapeName) {
    var drawing = sheet.getDrawingPatriarch();
    if (!(drawing instanceof XSSFDrawing xssfDrawing)) {
      throw new RuntimeException("シートからのDrawingの取得に失敗しました。");
    }

    // 図形を削除
    var ctDrawing = xssfDrawing.getCTDrawing();
    for (int i = 0; i < ctDrawing.getTwoCellAnchorList().size(); i++) {
      var anchor = ctDrawing.getTwoCellAnchorList().get(i);
      if (anchor.isSetSp() && anchor.getSp().getNvSpPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeTwoCellAnchor(i);
        return;
      }
      if (anchor.isSetGrpSp() && anchor.getGrpSp().getNvGrpSpPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeTwoCellAnchor(i);
        return;
      }
      if (anchor.isSetPic() && anchor.getPic().getNvPicPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeTwoCellAnchor(i);
        return;
      }
    }
    for (int i = 0; i < ctDrawing.getOneCellAnchorList().size(); i++) {
      var anchor = ctDrawing.getOneCellAnchorList().get(i);
      if (anchor.isSetSp() && anchor.getSp().getNvSpPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeOneCellAnchor(i);
        return;
      }
      if (anchor.isSetGrpSp() && anchor.getGrpSp().getNvGrpSpPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeOneCellAnchor(i);
        return;
      }
      if (anchor.isSetPic() && anchor.getPic().getNvPicPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeOneCellAnchor(i);
        return;
      }
    }
    for (int i = 0; i < ctDrawing.getAbsoluteAnchorList().size(); i++) {
      var anchor = ctDrawing.getAbsoluteAnchorList().get(i);
      if (anchor.isSetSp() && anchor.getSp().getNvSpPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeAbsoluteAnchor(i);
        return;
      }
      if (anchor.isSetGrpSp() && anchor.getGrpSp().getNvGrpSpPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeAbsoluteAnchor(i);
        return;
      }
      if (anchor.isSetPic() && anchor.getPic().getNvPicPr().getCNvPr().getName().equals(shapeName)) {
        ctDrawing.removeAbsoluteAnchor(i);
        return;
      }
    }
  }

}
