package com.qwerty0121.poi.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;

public class PoiSampleUtils {

  /**
   * テンプレートファイルからWorkbookを作成する
   * 
   * @param templateFileName テンプレートファイル名
   * @return Workbook
   * @throws IOException
   */
  public static Workbook createWorkbook(String templateFileName) throws IOException {
    try (
        var templateFileIS = PoiSampleUtils.class.getClassLoader()
            .getResourceAsStream(templateFileName);) {
      var workbook = WorkbookFactory.create(templateFileIS);
      return workbook;
    }
  }

  /**
   * Workbookを指定したファイル名を出力する
   * 
   * @param workbook Workbook
   * @throws IOException
   * @throws FileNotFoundException
   */
  public static void writeWorkbook(Workbook workbook, String fileName) throws IOException, FileNotFoundException {
    File outputFileDir = getOrCreateOutputDir();

    var outputFilePath = outputFileDir.toPath().resolve(fileName);
    try (var os = new FileOutputStream(outputFilePath.toString());) {
      workbook.write(os);
    }
  }

  /**
   * 画像ファイルをbyte配列として読み込む
   * 
   * @param fileName ファイル名
   * @return 画像ファイル(byte配列)
   * @throws FileNotFoundException
   * @throws IOException
   */
  public static byte[] loadPictureAsByteArray(String fileName) throws IOException {
    try (var is = PoiSampleUtils.class.getClassLoader().getResourceAsStream(Path.of("images", fileName).toString())) {
      return IOUtils.toByteArray(is);
    }
  }

  /**
   * 図形名をもとにシート内の図形を取得
   * 
   * @param sheet     シート
   * @param shapeName 図形名
   * @return 図形
   */
  public static XSSFShape getShapeByName(Sheet sheet, String shapeName) {
    if (!(sheet.getDrawingPatriarch() instanceof XSSFDrawing xssfDrawing)) {
      // XSSFDrawingでない場合は取得できないのでnullを返す
      return null;
    }

    return xssfDrawing.getShapes().stream()
        .filter(shape -> StringUtils.equals(shape.getShapeName(), shapeName))
        .findFirst()
        .orElse(null);
  }

  private static File getOrCreateOutputDir() {
    String outputFileDirPath = "./.output/";
    File outputFileDir = new File(outputFileDirPath);
    outputFileDir.mkdir();
    return outputFileDir;
  }

}
