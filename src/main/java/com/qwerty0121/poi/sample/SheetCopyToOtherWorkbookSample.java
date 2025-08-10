package com.qwerty0121.poi.sample;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qwerty0121.poi.utils.PoiSampleUtils;

/**
 * シートを別ワークブックにコピーするサンプル
 */
public class SheetCopyToOtherWorkbookSample {

  public static void main(String[] args) throws IOException {
    // コピー元ワークブックを読み込む
    var sourceWorkbook = (XSSFWorkbook) PoiSampleUtils.loadTemplateWorkbook("シートコピー(別ワークブック)テンプレート.xlsx");

    // ワークブックをコピーする
    try (var destinationWorkbook = copyWorkbook(sourceWorkbook)) {
      // コピーしたワークブックを出力する
      PoiSampleUtils.writeWorkbook(destinationWorkbook, "シートコピー(別ワークブック).xlsx");
    }
  }

  /**
   * ワークブックをコピーする
   * 
   * @param sourceWorkbook コピー元のワークブック
   * @return コピーしたワークブック
   */
  private static XSSFWorkbook copyWorkbook(XSSFWorkbook sourceWorkbook) {
    var newWorkbook = new XSSFWorkbook();

    // 全てのシートをコピーする
    sourceWorkbook.sheetIterator().forEachRemaining(sourceSheet -> {
      var newSheet = newWorkbook.createSheet(sourceSheet.getSheetName());
      copySheets((XSSFSheet) sourceSheet, newSheet);
    });

    return newWorkbook;
  }

  /**
   * シートをコピーする
   * 
   * @param sourceSheet コピー元のシート
   * @param newSheet    コピー先のシート
   */
  private static void copySheets(XSSFSheet sourceSheet, XSSFSheet newSheet) {
    // セルスタイルを保持するマップ
    var styleMap = new HashMap<CellStyle, CellStyle>();

    // シート内容の行をコピーする
    for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
      var sourceRow = sourceSheet.getRow(i);
      if (sourceRow == null) {
        continue; // 空の行はスキップ
      }

      // 行をコピー
      var newRow = newSheet.createRow(i);
      copyRow(sourceRow, newRow, styleMap);
    }

    // セル結合をコピー
    for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
      newSheet.addMergedRegion(sourceSheet.getMergedRegion(i));
    }
  }

  /**
   * 行をコピーする
   * 
   * @param sourceRow コピー元の行
   * @param newRow    コピー先の行
   * @param styleMap  セルスタイルを保持するマップ
   */
  private static void copyRow(Row sourceRow, Row newRow, Map<CellStyle, CellStyle> styleMap) {
    // 行内のセルをコピーする
    for (int i = sourceRow.getFirstCellNum(); i < sourceRow.getLastCellNum(); i++) {
      var sourceCell = sourceRow.getCell(i);
      if (sourceCell == null) {
        continue; // 空のセルはスキップ
      }

      // セルをコピー
      var newCell = newRow.createCell(i);
      copyCell(sourceCell, newCell, styleMap);
    }
  }

  /**
   * セルをコピーする
   * 
   * @param sourceCell コピー元のセル
   * @param newCell    コピー先のセル
   * @param styleMap   セルスタイルを保持するマップ
   */
  private static void copyCell(Cell sourceCell, Cell newCell, Map<CellStyle, CellStyle> styleMap) {
    // セルスタイルをコピー
    var sourceCellStyle = sourceCell.getCellStyle();
    if (sourceCellStyle != null) {
      var destinationCellStyle = styleMap.get(sourceCellStyle);
      if (destinationCellStyle == null) {
        // 新しいスタイルを作成してマップに保存
        destinationCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
        destinationCellStyle.cloneStyleFrom(sourceCellStyle);
        styleMap.put(sourceCellStyle, destinationCellStyle);
      }
      newCell.setCellStyle(destinationCellStyle);
    }

    // セルの値をコピー
    switch (sourceCell.getCellType()) {
      case STRING:
        newCell.setCellValue(sourceCell.getStringCellValue());
        break;
      case NUMERIC:
        newCell.setCellValue(sourceCell.getNumericCellValue());
        break;
      case BOOLEAN:
        newCell.setCellValue(sourceCell.getBooleanCellValue());
        break;
      case FORMULA:
        newCell.setCellFormula(sourceCell.getCellFormula());
        break;
      case BLANK:
        // ブランクセルは何もしない
        break;
      case ERROR:
        newCell.setCellErrorValue(sourceCell.getErrorCellValue());
        break;
      default:
        break;
    }

    // ハイパーリンクをコピー
    if (sourceCell.getHyperlink() != null) {
      var sourceLink = sourceCell.getHyperlink();
      var newLink = newCell.getSheet().getWorkbook().getCreationHelper().createHyperlink(sourceLink.getType());
      newLink.setAddress(sourceLink.getAddress());
      newLink.setLabel(sourceLink.getLabel());
      newCell.setHyperlink(newLink);
    }
  }

}
