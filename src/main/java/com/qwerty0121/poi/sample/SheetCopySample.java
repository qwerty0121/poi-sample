package com.qwerty0121.poi.sample;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * シートをコピーするサンプル
 */
public class SheetCopySample {

  public static void main(String[] args) throws Exception {
    var workbook = createWorkbook();

    // "テスト"シートを複製する
    var srcSheetIndex = workbook.getSheetIndex("テスト");
    workbook.cloneSheet(srcSheetIndex);

    writeWorkbook(workbook);
  }

  private static Workbook createWorkbook() throws IOException {
    try (var templateFileIS = SheetCopySample.class.getClassLoader().getResourceAsStream("シートコピーテンプレート.xls");) {
      var workbook = WorkbookFactory.create(templateFileIS);
      return workbook;
    }
  }

  private static void writeWorkbook(Workbook workbook) throws IOException, FileNotFoundException {
    File outputFileDir = getOrCreateDestDir();

    var outputFilePath = outputFileDir.toPath().resolve("シートコピー.xls");
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

}
