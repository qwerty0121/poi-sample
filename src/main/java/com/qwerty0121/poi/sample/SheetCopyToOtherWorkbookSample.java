package com.qwerty0121.poi.sample;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hssf.record.cf.PatternFormatting;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.Themes;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
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

    // シートの条件付き書式をコピー
    copySheetConditionalFormatting(sourceSheet.getSheetConditionalFormatting(),
        newSheet.getSheetConditionalFormatting(), sourceSheet.getWorkbook().getTheme());
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

  /**
   * シート条件付き書式をコピーする<br>
   * <br>
   * 条件種別が数式以外である場合はコピーされないので注意
   * 
   * @param srcSheetConditionalFormatting  コピー元のシート条件付き書式
   * @param destSheetConditionalFormatting コピー先のシート条件付き書式
   * @param srcTheme                       コピー元のテーマ
   */
  private static void copySheetConditionalFormatting(XSSFSheetConditionalFormatting srcSheetConditionalFormatting,
      XSSFSheetConditionalFormatting destSheetConditionalFormatting, Themes srcTheme) {
    // コピー元の条件付き書式における対象セル範囲とルールのペアをListとして取得
    var srcConditionalFormattingPairList = new ArrayList<Pair<CellRangeAddress[], XSSFConditionalFormattingRule>>();
    for (int i = 0; i < srcSheetConditionalFormatting.getNumConditionalFormattings(); i++) {
      var srcConditionalFormatting = srcSheetConditionalFormatting.getConditionalFormattingAt(i);
      var regions = srcConditionalFormatting.getFormattingRanges();
      for (int j = 0; j < srcConditionalFormatting.getNumberOfRules(); j++) {
        var srcRule = srcConditionalFormatting.getRule(j);
        srcConditionalFormattingPairList.add(Pair.of(regions, srcRule));
      }
    }

    // 条件付き書式の優先順位でソート
    srcConditionalFormattingPairList
        .sort(Comparator.comparing(conditionalFormattingPair -> conditionalFormattingPair.getValue().getPriority()));

    // 条件付き書式をコピーする
    srcConditionalFormattingPairList.forEach(srcConditionalFormattingPair -> {
      var regions = srcConditionalFormattingPair.getKey();
      var srcConditionalFormattingRule = srcConditionalFormattingPair.getValue();

      // コピー先の条件付き書式を作成
      XSSFConditionalFormattingRule destConditionalFormattingRule = null;
      if (srcConditionalFormattingRule.getConditionType() == ConditionType.FORMULA) {
        // 条件種別が数式である場合
        destConditionalFormattingRule = destSheetConditionalFormatting
            .createConditionalFormattingRule(srcConditionalFormattingRule.getFormula1());
      }

      if (destConditionalFormattingRule == null) {
        return;
      }

      // 条件付き書式ルールの書式設定をコピー
      copyConditionalFormattingRuleFormatting(srcConditionalFormattingRule, destConditionalFormattingRule, srcTheme);

      // コピーした条件付き書式を追加
      destSheetConditionalFormatting.addConditionalFormatting(regions, destConditionalFormattingRule);
    });
  }

  /**
   * 条件付き書式ルールにおける書式設定をコピーする<br>
   * <br>
   * NOTE: 以下の設定はコピーされないので注意
   * <ul>
   * <li>表示形式</li>
   * <li>フォント - 取り消し線</li>
   * </ul>
   * 
   * @param srcConditionalFormattingRule  コピー元の条件付き書式ルール
   * @param destConditionalFormattingRule コピー先の条件付き書式ルール
   * @param srcTheme                      コピー元のテーマ
   */
  private static void copyConditionalFormattingRuleFormatting(
      XSSFConditionalFormattingRule srcConditionalFormattingRule,
      XSSFConditionalFormattingRule destConditionalFormattingRule, Themes srcTheme) {
    if (srcConditionalFormattingRule.getFontFormatting() != null) {
      // フォント設定をコピー
      var srcFontFormatting = srcConditionalFormattingRule.getFontFormatting();
      var destFontFormatting = destConditionalFormattingRule.createFontFormatting();

      // スタイル
      destFontFormatting.setFontStyle(srcFontFormatting.isItalic(), srcFontFormatting.isBold());
      // サイズ
      if (srcFontFormatting.getFontHeight() != -1) {
        destFontFormatting.setFontHeight(srcFontFormatting.getFontHeight());
      }
      // 下線
      destFontFormatting.setUnderlineType(srcFontFormatting.getUnderlineType());
      // 色
      var destFontColor = cloneColorFrom(srcFontFormatting.getFontColor(), srcTheme);
      if (destFontColor != null) {
        destFontFormatting.setFontColor(destFontColor);
      }
    }

    if (srcConditionalFormattingRule.getBorderFormatting() != null) {
      // 罫線設定をコピー
      var srcBorderFormatting = srcConditionalFormattingRule.getBorderFormatting();
      var destBorderFormatting = destConditionalFormattingRule.createBorderFormatting();

      // 線のスタイル
      destBorderFormatting.setBorderBottom(srcBorderFormatting.getBorderBottom());
      destBorderFormatting.setBorderLeft(srcBorderFormatting.getBorderLeft());
      destBorderFormatting.setBorderRight(srcBorderFormatting.getBorderRight());
      destBorderFormatting.setBorderTop(srcBorderFormatting.getBorderTop());

      // 線の色
      var destBottomColor = cloneColorFrom(srcBorderFormatting.getBottomBorderColorColor(), srcTheme);
      if (destBottomColor != null) {
        destBorderFormatting.setBottomBorderColor(destBottomColor);
      }
      var destLeftColor = cloneColorFrom(srcBorderFormatting.getLeftBorderColorColor(), srcTheme);
      if (destLeftColor != null) {
        destBorderFormatting.setLeftBorderColor(destLeftColor);
      }
      var destRightColor = cloneColorFrom(srcBorderFormatting.getRightBorderColorColor(), srcTheme);
      if (destRightColor != null) {
        destBorderFormatting.setRightBorderColor(destRightColor);
      }
      var destTopColor = cloneColorFrom(srcBorderFormatting.getTopBorderColorColor(), srcTheme);
      if (destTopColor != null) {
        destBorderFormatting.setTopBorderColor(destTopColor);
      }
    }

    if (srcConditionalFormattingRule.getPatternFormatting() != null) {
      // 塗りつぶし設定をコピー
      var srcPatternFormatting = srcConditionalFormattingRule.getPatternFormatting();
      var destPatternFormatting = destConditionalFormattingRule.createPatternFormatting();

      // 背景色
      var destFillBackgroundColor = cloneColorFrom(srcPatternFormatting.getFillBackgroundColorColor(), srcTheme);
      if (destFillBackgroundColor != null) {
        destPatternFormatting.setFillBackgroundColor(destFillBackgroundColor);
      }
      // パターンの色
      var destFillForegroundColor = cloneColorFrom(srcPatternFormatting.getFillForegroundColorColor(), srcTheme);
      if (destFillForegroundColor != null) {
        destPatternFormatting.setFillForegroundColor(destFillForegroundColor);
      }
      // パターンスタイル
      if (srcPatternFormatting.getFillPattern() != PatternFormatting.NO_FILL) {
        destPatternFormatting.setFillPattern(srcPatternFormatting.getFillPattern());
      }
    }
  }

  /**
   * 色を複製する
   * 
   * @param srcColor 複製元の色
   * @param theme    複製元のテーマ
   * @return 複製した色。複製元がnullの場合はnullを返す。
   */
  private static XSSFColor cloneColorFrom(XSSFColor srcColor, Themes theme) {
    if (srcColor == null) {
      return null;
    }

    var newColor = XSSFColor.from(srcColor.getCTColor());

    if (srcColor.isThemed()) {
      // テーマカラーの場合、テーマから取得したRGB値を設定
      newColor
          .setRGB(Optional.ofNullable(theme.getThemeColor(srcColor.getTheme())).map(XSSFColor::getRGB).orElse(null));
    }

    return newColor;
  }

}
