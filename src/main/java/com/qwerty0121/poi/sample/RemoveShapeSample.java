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
      removeSimpleShape(sheet);

      // 図形グループを削除する
      removeShapeGroup(sheet);

      // 画像を削除する
      removePicture(sheet);

      PoiSampleUtils.writeWorkbook(workbook, "図形削除サンプル.xlsx");
    }
  }

  /**
   * 図形を削除する
   * 
   * @param sheet シート
   * @throws IOException
   */
  private static void removeSimpleShape(Sheet sheet) throws IOException {
    var drawing = sheet.getDrawingPatriarch();
    if (!(drawing instanceof XSSFDrawing xssfDrawing)) {
      throw new RuntimeException("シートからのDrawingの取得に失敗しました。");
    }

    // 削除対象の図形の名前
    var name = "shape";

    // 図形を削除
    xssfDrawing.getCTDrawing().getOneCellAnchorList()
        .removeIf(anchor -> anchor.isSetSp() && anchor.getSp().getNvSpPr().getCNvPr().getName().equals(name));
    xssfDrawing.getCTDrawing().getTwoCellAnchorList()
        .removeIf(anchor -> anchor.isSetSp() && anchor.getSp().getNvSpPr().getCNvPr().getName().equals(name));
    xssfDrawing.getCTDrawing().getAbsoluteAnchorList()
        .removeIf(anchor -> anchor.isSetSp() && anchor.getSp().getNvSpPr().getCNvPr().getName().equals(name));
  }

  /**
   * 図形グループを削除する
   * 
   * @param sheet シート
   * @throws IOException
   */
  private static void removeShapeGroup(Sheet sheet) throws IOException {
    var drawing = sheet.getDrawingPatriarch();
    if (!(drawing instanceof XSSFDrawing xssfDrawing)) {
      throw new RuntimeException("シートからのDrawingの取得に失敗しました。");
    }

    // 削除対象の図形グループの名前
    var name = "shape-group";

    // 図形グループを削除
    xssfDrawing.getCTDrawing().getOneCellAnchorList()
        .removeIf(anchor -> anchor.isSetGrpSp() && anchor.getGrpSp().getNvGrpSpPr().getCNvPr().getName().equals(name));
    xssfDrawing.getCTDrawing().getTwoCellAnchorList()
        .removeIf(anchor -> anchor.isSetGrpSp() && anchor.getGrpSp().getNvGrpSpPr().getCNvPr().getName().equals(name));
    xssfDrawing.getCTDrawing().getAbsoluteAnchorList()
        .removeIf(anchor -> anchor.isSetGrpSp() && anchor.getGrpSp().getNvGrpSpPr().getCNvPr().getName().equals(name));
  }

  /**
   * 画像を削除する
   * 
   * @param sheet シート
   * @throws IOException
   */
  private static void removePicture(Sheet sheet) throws IOException {
    var drawing = sheet.getDrawingPatriarch();
    if (!(drawing instanceof XSSFDrawing xssfDrawing)) {
      throw new RuntimeException("シートからのDrawingの取得に失敗しました。");
    }

    // 削除対象の画像の名前
    var name = "picture";

    // 画像を削除
    xssfDrawing.getCTDrawing().getOneCellAnchorList()
        .removeIf(anchor -> anchor.isSetPic() && anchor.getPic().getNvPicPr().getCNvPr().getName().equals(name));
    xssfDrawing.getCTDrawing().getTwoCellAnchorList()
        .removeIf(anchor -> anchor.isSetPic() && anchor.getPic().getNvPicPr().getCNvPr().getName().equals(name));
    xssfDrawing.getCTDrawing().getAbsoluteAnchorList()
        .removeIf(anchor -> anchor.isSetPic() && anchor.getPic().getNvPicPr().getCNvPr().getName().equals(name));
  }

}
