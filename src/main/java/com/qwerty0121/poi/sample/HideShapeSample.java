package com.qwerty0121.poi.sample;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

import com.qwerty0121.poi.utils.PoiSampleUtils;

/**
 * 図形を非表示にするサンプル
 */
public class HideShapeSample {

  public static void main(String[] args) throws IOException {
    try (var workbook = PoiSampleUtils.loadTemplateWorkbook("図形非表示サンプルテンプレート.xlsx");) {
      // "テスト"シートを取得
      var sheet = workbook.getSheet("テスト");

      // 画像を非表示にする
      hidePicture(sheet);

      // 図形を非表示にする
      hideSimpleShape(sheet);

      // 図形グループを非表示にする
      hideShapeGroup(sheet);

      PoiSampleUtils.writeWorkbook(workbook, "図形非表示サンプル.xlsx");
    }
  }

  /**
   * 画像を非表示にする
   * 
   * @param sheet シート
   * @throws IOException
   */
  private static void hidePicture(Sheet sheet) throws IOException {
    // 画像図形を取得
    var shape = PoiSampleUtils.getShapeByName(sheet, "picture");
    if (!(shape instanceof XSSFPicture picture)) {
      throw new RuntimeException("テンプレートファイルに「picture」という名前の画像図形が存在しません。");
    }

    // 画像図形を非表示にする
    picture.getCTPicture().getNvPicPr().getCNvPr().setHidden(true);
  }

  /**
   * 図形を非表示にする
   * 
   * @param sheet シート
   * @throws IOException
   */
  private static void hideSimpleShape(Sheet sheet) throws IOException {
    // 図形を取得
    var shape = PoiSampleUtils.getShapeByName(sheet, "shape");
    if (!(shape instanceof XSSFSimpleShape simpleShape)) {
      throw new RuntimeException("テンプレートファイルに「shape」という名前の図形が存在しません。");
    }

    // 図形を非表示にする
    simpleShape.getCTShape().getNvSpPr().getCNvPr().setHidden(true);
  }

  /**
   * 図形グループを非表示にする
   * 
   * @param sheet シート
   * @throws IOException
   */
  private static void hideShapeGroup(Sheet sheet) throws IOException {
    // 図形グループを取得
    var shape = PoiSampleUtils.getShapeByName(sheet, "shape-group");
    if (!(shape instanceof XSSFShapeGroup shapeGroup)) {
      throw new RuntimeException("テンプレートファイルに「shape-group」という名前の図形グループが存在しません。");
    }

    // 図形グループを非表示にする
    shapeGroup.getCTGroupShape().getNvGrpSpPr().getCNvPr().setHidden(true);
  }

}
