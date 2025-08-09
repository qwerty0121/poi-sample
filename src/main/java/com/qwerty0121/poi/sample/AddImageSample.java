package com.qwerty0121.poi.sample;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import com.qwerty0121.poi.utils.PoiSampleUtils;

/**
 * 画像を追加するサンプル
 */
public class AddImageSample {

  public static void main(String[] args) throws IOException {
    try (var workbook = PoiSampleUtils.loadTemplateWorkbook("画像追加サンプルテンプレート.xlsx");) {
      // "テスト"シートを取得
      var sheet = workbook.getSheet("テスト");

      // 画像をシートに追加
      addImage(workbook, sheet);

      // テンプレートの図形に画像を追加
      addImageToContainer(workbook, sheet);

      PoiSampleUtils.writeWorkbook(workbook, "画像追加サンプル.xlsx");
    }
  }

  private static void addImage(Workbook workbook, Sheet sheet) throws IOException {
    // 画像ファイルを読み込み
    var image = PoiSampleUtils.loadPictureAsByteArray("add-image-sample.png");
    var imageIdx = workbook.addPicture(image, Workbook.PICTURE_TYPE_PNG);

    // 画像の追加位置とサイズを指定
    var anchor = workbook.getCreationHelper().createClientAnchor();
    // 1. 画像の追加位置となるセル範囲を指定
    anchor.setCol1(1);
    anchor.setRow1(3);
    anchor.setCol2(11);
    anchor.setRow2(13);
    // 2. 1で指定したセル範囲の左上/右下からの座標を指定する
    // ※ここで指定した左上/右下の座標が画像が追加される座標となる。
    anchor.setDx1(Units.EMU_PER_PIXEL * 10);
    anchor.setDy1(Units.EMU_PER_PIXEL * 10);
    anchor.setDx2(Units.EMU_PER_PIXEL * -10);
    anchor.setDy2(Units.EMU_PER_PIXEL * -10);

    // シートに画像を追加
    var patriarch = sheet.createDrawingPatriarch();
    patriarch.createPicture(anchor, imageIdx);
  }

  private static void addImageToContainer(Workbook workbook, Sheet sheet) throws IOException {
    // 画像ファイルを読み込み
    var image = PoiSampleUtils.loadPictureAsByteArray("add-image-sample.png");
    var imageIdx = workbook.addPicture(image, Workbook.PICTURE_TYPE_PNG);

    // 画像の追加位置の目安となる図形
    var container = PoiSampleUtils.getShapeByName(sheet, "image-container");
    var containerAnchor = (XSSFClientAnchor) container.getAnchor();

    // 図形と同じ位置に画像を追加
    var patriarch = sheet.createDrawingPatriarch();
    var anchor = patriarch.createAnchor(
        containerAnchor.getDx1(),
        containerAnchor.getDy1(),
        containerAnchor.getDx2(),
        containerAnchor.getDy2(),
        containerAnchor.getCol1(),
        containerAnchor.getRow1(),
        containerAnchor.getCol2(),
        containerAnchor.getRow2());
    patriarch.createPicture(anchor, imageIdx);
  }

}
