package com.qwerty0121.poi.sample;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

/**
 * 図形にテキストを設定するサンプル
 */
public class ShapeTextSettingSample {

  public static void main(String[] args) throws Exception {
    var workbook = createWorkbook();

    // "テスト"シートを取得
    var sheet = workbook.getSheet("テスト");

    // 図形にテキストを設定
    replaceShapeText(sheet, "設定対象図形", "${text}", "プログラムから設定したテキスト");

    writeWorkbook(workbook);
  }

  private static Workbook createWorkbook() throws IOException {
    try (
        var templateFileIS = ShapeTextSettingSample.class.getClassLoader()
            .getResourceAsStream("図形テキスト設定テンプレート.xlsx");) {
      var workbook = WorkbookFactory.create(templateFileIS);
      return workbook;
    }
  }

  private static void writeWorkbook(Workbook workbook) throws IOException, FileNotFoundException {
    File outputFileDir = getOrCreateDestDir();

    var outputFilePath = outputFileDir.toPath().resolve("図形テキスト設定.xlsx");
    try (var os = new FileOutputStream(outputFilePath.toString());) {
      workbook.write(os);
    }
  }

  private static File getOrCreateDestDir() {
    String outputFileDirPath = "./.output/";
    File outputFileDir = new File(outputFileDirPath);
    outputFileDir.mkdir();
    return outputFileDir;
  }

  /**
   * 図形のテキストを置換
   * NOTE: TextRunごとに置換処理を行う。
   * 
   * @param sheet        シート
   * @param shapeName    テキストを設定する図形の名前
   * @param searchString 置換対象の文字列
   * @param replacement  置換後の文字列
   */
  private static void replaceShapeText(Sheet sheet, String shapeName, String searchString, String replacement) {
    var targetShape = getShapeByName(sheet, shapeName);
    if (targetShape == null) {
      // 図形名から図形を取得できなかった場合は何もしない
      return;
    }

    if (!(targetShape instanceof XSSFSimpleShape xssfSimpleShape)) {
      // 図形がXSSFSimpleShapeでない場合は何もしない
      return;
    }

    xssfSimpleShape.getTextParagraphs().forEach(textParagraph -> {
      textParagraph.getTextRuns().stream().forEach(textRun -> {
        // 元々図形に設定されているテキストを取得
        var originalText = textRun.getText();

        // 置換対象のタグを置換する
        var replacedText = StringUtils.replace(originalText, searchString, replacement);

        // 図形にテキストを設定
        textRun.setText(replacedText);
      });
    });
  }

  /**
   * 図形名をもとにシート内の図形を取得
   * 
   * @param sheet     シート
   * @param shapeName 図形名
   * @return 図形
   */
  private static XSSFShape getShapeByName(Sheet sheet, String shapeName) {
    if (!(sheet.getDrawingPatriarch() instanceof XSSFDrawing xssfDrawing)) {
      // XSSFDrawingでない場合は取得できないのでnullを返す
      return null;
    }

    return xssfDrawing.getShapes().stream()
        .filter(shape -> StringUtils.equals(shape.getShapeName(), shapeName))
        .findFirst()
        .orElse(null);
  }

}
