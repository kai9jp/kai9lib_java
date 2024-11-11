package kai9.libs;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.stream.IntStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * POI（エクセル操作）の自作ユーティリティ
 */
public class PoiUtil {

    /**
     * セルの値を返す
     */
    public static String GetCellValue(Sheet sheet, Integer Row, Integer Col) {
        Row row = sheet.getRow(Row);
        Cell cell = row.getCell(Col);
        return cell.getStringCellValue();
    }

    /**
     * 文字列検索し、最初に発見した行を返す
     */
    public static int findRow(Sheet sheet, String cellContent) {
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equals(cellContent)) {
                        return row.getRowNum();
                    }
                }
            }
        }
        return -1;
    }

    /**
     * 文字列検索し、最初に発見した列を返す
     */
    public static int findCol(Sheet sheet, String cellContent) {
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equals(cellContent)) {
                        return cell.getColumnIndex();
                    }
                }
            }
        }
        return -1;
    }

    public static int findCol(Sheet sheet, Integer Row, String cellContent) {
        if (Row == -1) return -1;
        Row row = sheet.getRow(Row);
        for (Cell cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1;
    }

    /**
     * 文字列検索し、最初に発見したセルを返す
     */
    public static Cell findCell(Sheet sheet, String cellContent) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                        return cell;
                    }
                }
            }
        }
        return null;
    }

    /**
     * セルの文字を取得
     */
    public static String GetStringValue(Sheet sheet, Integer Row, Integer Col) {
        Row row = sheet.getRow(Row);
        if (row == null) {
            return "";
        } // 空行の場合nullが返る
        Cell cell = row.getCell(Col);
        if (cell == null) {
            return "";
        } // 空の場合nullが返る

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cell.getCellType() == CellType.NUMERIC) {
//        	return String.valueOf(cell.getNumericCellValue());
            DataFormatter dataFormatter = new DataFormatter();
            return dataFormatter.formatCellValue(cell);
        }
        if (cell.getCellType() == CellType.FORMULA) {
            Workbook book = cell.getSheet().getWorkbook();
            CreationHelper helper = book.getCreationHelper();
            FormulaEvaluator evaluator = helper.createFormulaEvaluator();
            CellValue value = evaluator.evaluate(cell);
            if (value.getCellType() == CellType.STRING) {
                return value.getStringValue();
            }
            if (value.getCellType() == CellType.NUMERIC) {
//                return  String.valueOf(value.getNumberValue());        
                DataFormatter dataFormatter = new DataFormatter();
                return dataFormatter.formatCellValue(cell, evaluator);
            }
        }
        return "";
    }

    /**
     * セルの文字を取得(bool型)
     */
    public static boolean GetBoolValue(Sheet sheet, Integer Row, Integer Col) {
        Row row = sheet.getRow(Row);
        Cell cell = row.getCell(Col);
        try {
            return cell.getBooleanCellValue();
        } catch (Exception e) {
            // フラグ値以外だと例外が発生するのでfalseを返す
            return false;
        }
    }

    /**
     * セルの文字を取得(数値型)
     */
    public static Number GetNumericValue(Sheet sheet, Integer Row, Integer Col) {
        Row row = sheet.getRow(Row);
        Cell cell = row.getCell(Col);
        try {
            return cell.getNumericCellValue();
        } catch (Exception e) {
            // 数値以外だと例外が発生するのでゼロを返す
            return 0;
        }
    }

    // 指定された行を削除する
    public static void removeRow(Sheet sheet, int rowIndex) {
        // シートの最後の行番号を取得
        int lastRowNum = sheet.getLastRowNum();

        // poiのバグで、いずれ直るかもしれないが、5.2.3や5.3.0では、シフトする範囲に、エクセルの名前を参照しているセルが有る場合、
        // 下記の例外が発生してしまうため、名前を退避し、後で戻すという事で回避している。
        // java.lang.IllegalArgumentException: Cell reference invalid: #REF!
        // at org.apache.poi.ss.util.CellReference.(CellReference.java:112)

        // ワークブック内の全ての名前を取得し、List<Name> にキャスト
        List<? extends Name> allNamesWildcard = sheet.getWorkbook().getAllNames();
        @SuppressWarnings("unchecked")
        List<Name> allNames = (List<Name>) allNamesWildcard;

        // 名前とダミー文字列のマップを作成
        Map<String, String> nameToDummyMap = new HashMap<>();
        for (Name name : allNames) {
            nameToDummyMap.put(name.getNameName(), "\"" + name.getNameName() + "【計算式ダミー】\"");
        }

        // シフト対象のセルが計算式かどうかを確認し、ダミー文字列を追加
        for (int rowNum = rowIndex + 1; rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.FORMULA) {
                        String formula = cell.getCellFormula();
                        for (String name : nameToDummyMap.keySet()) {
                            if (formula.contains(name)) {
                                // 計算式内の名前にダミー文字列を追加
                                formula = formula.replace(name, nameToDummyMap.get(name));
                            }
                        }
                        cell.setCellFormula(formula);
                    }
                }
            }
        }
        // 削除する行がシートの範囲内にある場合
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            // 指定された行より下の行が存在するか確認
            if (sheet.getRow(rowIndex + 1) != null) {
                // 指定された行より下の行を上にシフトして、行を削除
                sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
            }
        }
        // 削除行がシート最終行の場合
        else if (rowIndex == lastRowNum) {
            // 削除する行を取得
            Row removingRow = sheet.getRow(rowIndex);
            // 行が存在する場合、その行を削除
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }

        // 名前の定義を元に戻す
        for (int rowNum = rowIndex; rowNum < lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.FORMULA) {
                        String formula = cell.getCellFormula();
                        for (String dummyName : nameToDummyMap.values()) {
                            if (formula.contains(dummyName)) {
                                // ダミー文字列を元の名前に戻す
                                String originalName = dummyName.replace("\"", "").replace("【計算式ダミー】", "");
                                formula = formula.replace(dummyName, originalName);
                            }
                        }
                        cell.setCellFormula(formula);
                    }
                }
            }
        }

    }

    // 指定された行を削除する(範囲指定で高速化)
    public static void removeRows(Sheet sheet, List<Integer> rowIndexes) {
        // ワークブック内の全ての名前を取得し、List<Name> にキャスト
        List<? extends Name> allNamesWildcard = sheet.getWorkbook().getAllNames();
        @SuppressWarnings("unchecked")
        List<Name> allNames = (List<Name>) allNamesWildcard;

        // 名前とダミー文字列のマップを作成
        Map<String, String> nameToDummyMap = new HashMap<>();
        for (Name name : allNames) {
            nameToDummyMap.put(name.getNameName(), "\"" + name.getNameName() + "【計算式ダミー】\"");
        }

        // シフト対象のセルが計算式かどうかを確認し、ダミー文字列を追加
        for (int rowNum : rowIndexes) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.FORMULA) {
                        String formula = cell.getCellFormula();
                        if (!formula.contains("【計算式ダミー】")) {// 既に置換済の場合は処理しない(何故か置換済の物が混入する事象の対策)
                            for (String name : nameToDummyMap.keySet()) {
                                if (formula.contains(name)) {
                                    // 計算式内の名前にダミー文字列を追加
                                    formula = formula.replace(name, nameToDummyMap.get(name));
                                }
                            }
                            cell.setCellFormula(formula);
                        }
                    }
                }
            }
        }

        // 行インデックスを昇順にソート
        Collections.sort(rowIndexes);

        // 削除する行の連続範囲を保持するリスト
        List<int[]> ranges = new ArrayList<>();
        int start = rowIndexes.get(0);
        int end = start;

        // 連続範囲を計算
        for (int i = 1; i < rowIndexes.size(); i++) {
            if (rowIndexes.get(i) == end + 1) {
                // 次の行も連続範囲に含まれる場合
                end = rowIndexes.get(i);
            } else {
                // 連続範囲の終了をリストに追加し、新しい範囲を開始
                ranges.add(new int[] { start, end });
                start = rowIndexes.get(i);
                end = start;
            }
        }
        // 最後の連続範囲をリストに追加
        ranges.add(new int[] { start, end });

        int totalShift = 0; // 全体のシフトカウント
        int lastRowNum = sheet.getLastRowNum(); // シートの最後の行番号

        // 各連続範囲についてシフト操作を行う
        for (int[] range : ranges) {
            int rangeStart = range[0] - totalShift; // シフト開始位置
            int rangeEnd = range[1] - totalShift; // シフト終了位置

            // シフト範囲がシートの最後の行を超えないように調整
            if (rangeEnd > lastRowNum) {
                rangeEnd = lastRowNum;
            }

            // シフトする行数を計算
            int shiftCount;
            if (rangeEnd < lastRowNum) {
                shiftCount = rangeEnd - rangeStart + 1; // シフトする行数
            } else {
                shiftCount = rangeEnd - rangeStart; // lastRowNumに達した場合は+1をしない
            }

            // シフト対象が範囲外の場合は無視
            if (rangeStart + shiftCount > lastRowNum) continue;

            // シフト操作を実行
            sheet.shiftRows(rangeStart + shiftCount, lastRowNum, -shiftCount);
            totalShift += shiftCount;
        }

        // シフト後の空行を削除
        for (int i = 0; i < totalShift; i++) {
            Row row = sheet.getRow(lastRowNum - i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }

        // 名前の定義を元に戻す
        for (int rowNum : rowIndexes) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.FORMULA) {
                        String formula = cell.getCellFormula();
                        for (String dummyName : nameToDummyMap.values()) {
                            if (formula.contains(dummyName)) {
                                // ダミー文字列を元の名前に戻す
                                String originalName = dummyName.replace("\"", "").replace("【計算式ダミー】", "");
                                formula = formula.replace(dummyName, originalName);
                            }
                        }
                        cell.setCellFormula(formula);
                    }
                }
            }
        }
    }

    // 指定されたインデックスに行を挿入
    public static void insertRow(Sheet sheet, int rowIndex) {
        // シートの最後の行番号を取得
        int lastRowNum = sheet.getLastRowNum();
        // 挿入する行がシートの範囲内にある場合
        if (rowIndex >= 0 && rowIndex <= lastRowNum) {
            // 指定された行から最後の行までを下にシフトして、行を挿入
            sheet.shiftRows(rowIndex, lastRowNum, 1);
        }
        // 新しい行を指定されたインデックスに作成
        sheet.createRow(rowIndex);
    }

    // 指定されたインデックスの行の書式をコピー
    public static void copyRowFormatting(Sheet sheet, int sourceRowIndex, int targetRowIndex) {
        // ソース行とターゲット行を取得
        Row sourceRow = sheet.getRow(sourceRowIndex);
        Row targetRow = sheet.getRow(targetRowIndex);

        // ソース行とターゲット行の両方が存在する場合
        if (sourceRow != null && targetRow != null) {
            // ソース行のすべてのセルを走査
            for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                // ソースセルを取得
                Cell sourceCell = sourceRow.getCell(i);
                if (sourceCell != null) {
                    // ターゲットセルを作成
                    Cell targetCell = targetRow.createCell(i);
                    // ソースセルの書式をターゲットセルにコピー
                    copyCellFormatting(sourceCell, targetCell);
                }
            }
        }
    }

    // 指定されたセルの書式をコピー
    private static void copyCellFormatting(Cell sourceCell, Cell targetCell) {
        // スタイルをコピー
        CellStyle newCellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);

        // コメントをコピー（ある場合）
        if (sourceCell.getCellComment() != null) {
            targetCell.setCellComment(sourceCell.getCellComment());
        }

        // セルの内容をコピー（セルタイプに応じて）
        switch (sourceCell.getCellType()) {
        case STRING:
            targetCell.setCellValue(sourceCell.getStringCellValue());
            break;
        case NUMERIC:
            targetCell.setCellValue(sourceCell.getNumericCellValue());
            break;
        case BOOLEAN:
            targetCell.setCellValue(sourceCell.getBooleanCellValue());
            break;
        case FORMULA:
            targetCell.setCellFormula(sourceCell.getCellFormula());
            break;
        case BLANK:
            targetCell.setBlank();
            break;
        default:
            break;
        }
    }

    /**
     * 全てのセルのフォントを指定されたフォント名とフォントサイズに設定
     *
     * @param sheet シート
     * @param fontName フォント名
     * @param fontSize フォントサイズ
     */
    public static void setFont(Sheet sheet, String fontName, int fontSize) {
        Workbook workbook = sheet.getWorkbook();

        // 新しいフォントを作成
        Font newFont = workbook.createFont();
        newFont.setFontName(fontName);
        newFont.setFontHeightInPoints((short) fontSize);

        // デフォルトのスタイルを作成
        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(newFont);

        // スタイルのキャッシュを作成
        Map<CellStyle, CellStyle> styleCache = new HashMap<>();

        for (Row row : sheet) {
            for (Cell cell : row) {
                CellStyle existingStyle = cell.getCellStyle();
                if (existingStyle != null) {
                    // 既存のスタイルがキャッシュにあるか確認
                    CellStyle cachedStyle = styleCache.get(existingStyle);
                    if (cachedStyle == null) {
                        // キャッシュにない場合、新しいスタイルを作成してキャッシュに追加
                        CellStyle newStyle = workbook.createCellStyle();
                        newStyle.cloneStyleFrom(existingStyle);
                        newStyle.setFont(newFont);
                        styleCache.put(existingStyle, newStyle);
                        cachedStyle = newStyle;
                    }
                    // セルにキャッシュされたスタイルを設定
                    cell.setCellStyle(cachedStyle);
                } else {
                    // 既存のスタイルがない場合、新しいスタイルを直接設定
                    cell.setCellStyle(baseStyle);
                }
            }
        }
    }

    public static void setCellFormatAsText(Sheet sheet, CellRangeAddress range) {
        Workbook workbook = sheet.getWorkbook();

        // スタイルのキャッシュを作成
        Map<CellStyle, CellStyle> styleCache = new HashMap<>();

        DataFormat dataFormat = workbook.createDataFormat();
        CellStyle newStyle;

        for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }

                CellStyle existingStyle = cell.getCellStyle();

                // 既存のスタイルがキャッシュにあるか確認
                newStyle = styleCache.get(existingStyle);
                if (newStyle == null) {
                    newStyle = workbook.createCellStyle();
                    if (existingStyle != null) {
                        newStyle.cloneStyleFrom(existingStyle);
                    }

                    // セル書式を文字列に設定
                    newStyle.setDataFormat(dataFormat.getFormat("@"));

                    styleCache.put(existingStyle, newStyle);
                }

                // セルに新しいスタイルを設定
                cell.setCellStyle(newStyle);
            }
        }
    }

    /**
     * シートの範囲内のセルに背景色とフォント色を設定
     * 
     * @param sheet シート
     * @param range 範囲を表すCellRangeAddress
     * @param bgColor 背景色
     * @param fontColor フォント色
     * 
     *        使用例 ↓
     * 
     *        範囲を指定する
     *        CellRangeAddress range = new CellRangeAddress(2, 5, 3, 7); ※startRow, endRow, startCol, endCol
     *
     *        セルに背景色とフォント色を設定する
     *        setCellBackgroundAndFontColor(sheet, range, IndexedColors.LIGHT_GREEN, IndexedColors.WHITE);
     * 
     */
    // IndexedColorsを使用するメソッド
    public static CellStyle setCellBackgroundAndFontColor(Sheet sheet, CellRangeAddress range, Object bgColor, Object fontColor, CellStyle baseStyle) {
        CellStyle newStyle = baseStyle;

        if (newStyle == null) {
            Workbook workbook = sheet.getWorkbook();

            // スタイルのキャッシュを作成
            Map<CellStyle, CellStyle> styleCache = new HashMap<>();
            for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = sheet.createRow(rowIndex);
                }

                for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell(colIndex);
                    }

                    CellStyle existingStyle = cell.getCellStyle();
                    Font existingFont = null;

                    // 既存のスタイルがあり、それがキャッシュされているか確認
                    if (existingStyle != null) {
                        existingFont = workbook.getFontAt(existingStyle.getFontIndex());
                        if (styleCache.containsKey(existingStyle)) {
                            newStyle = styleCache.get(existingStyle);
                        } else {
                            boolean backgroundMatches = false;
                            boolean fontColorMatches = false;

                            // 背景色が一致するか確認
                            if (bgColor instanceof IndexedColors) {
                                short bgColorIndex = ((IndexedColors) bgColor).getIndex();
                                short existingColorIndex = existingStyle.getFillForegroundColor();

                                // 自動色 (64) は白色 (9) として扱う
                                if (existingColorIndex == 64 && bgColorIndex == IndexedColors.WHITE.getIndex()) {
                                    backgroundMatches = true;
                                } else {
                                    backgroundMatches = existingColorIndex == bgColorIndex;
                                }
                            } else if (bgColor instanceof XSSFColor) {
                                if (existingStyle.getFillForegroundColorColor() != null) {
                                    backgroundMatches = existingStyle.getFillForegroundColorColor().equals(bgColor);
                                }
                            }

                            // フォント色が一致するか確認
                            if (existingFont != null) {
                                short fontColorIndex = existingFont.getColor();
                                short desiredColorIndex = (fontColor instanceof IndexedColors) ? ((IndexedColors) fontColor).getIndex() : -1;

                                // 自動色 (0) は黒色 (IndexedColors.BLACK.getIndex()) として扱う
                                if ((fontColorIndex == 0 || fontColorIndex == IndexedColors.AUTOMATIC.getIndex()) && desiredColorIndex == IndexedColors.BLACK.getIndex()) {
                                    fontColorMatches = true;
                                } else {
                                    fontColorMatches = fontColorIndex == desiredColorIndex;
                                }
                            }

                            // 背景色とフォント色が既存のスタイルと一致する場合、そのまま流用
                            if (backgroundMatches && fontColorMatches) {
                                newStyle = existingStyle;
                            }
                        }
                    }

                    if (newStyle == null) {
                        // 一致しない場合、新しいスタイルを作成
                        newStyle = workbook.createCellStyle();
                        if (existingStyle != null) {
                            newStyle.cloneStyleFrom(existingStyle);
                        }

                        // 背景色を設定
                        if (bgColor instanceof IndexedColors) {
                            newStyle.setFillForegroundColor(((IndexedColors) bgColor).getIndex());
                        } else if (bgColor instanceof XSSFColor) {
                            newStyle.setFillForegroundColor((XSSFColor) bgColor);
                        }
                        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                        // フォント色を設定
                        XSSFFont newFont = (XSSFFont) workbook.createFont();
                        if (existingFont != null) {
                            newFont.setFontName(existingFont.getFontName());
                            newFont.setFontHeightInPoints(existingFont.getFontHeightInPoints());
                            newFont.setBold(existingFont.getBold());
                            newFont.setItalic(existingFont.getItalic());
                            newFont.setStrikeout(existingFont.getStrikeout());
                            newFont.setTypeOffset(existingFont.getTypeOffset());
                            newFont.setUnderline(existingFont.getUnderline());
                        }

                        if (fontColor instanceof IndexedColors) {
                            newFont.setColor(((IndexedColors) fontColor).getIndex());
                        } else if (fontColor instanceof XSSFColor) {
                            throw new UnsupportedOperationException("XSSFColorの設定はバグがあり上手く機能しないので利用不可");
                        }

                        newStyle.setFont(newFont);
                        styleCache.put(existingStyle, newStyle);
                    }

                    // セルに新しいスタイルを設定
                    cell.setCellStyle(newStyle);
                }
            }
        } else {
            // 基本スタイルが渡されていた場合、そのまま使用する
//	        for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
//	            Row row = sheet.getRow(rowIndex);
//	            if (row == null) {
//	                row = sheet.createRow(rowIndex);
//	            }
//
//	            for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
//	                Cell cell = row.getCell(colIndex);
//	                if (cell == null) {
//	                    cell = row.createCell(colIndex);
//	                }
//
//	                // セルに既存のスタイルを設定
//	                cell.setCellStyle(newStyle);
//	            }
//	        }

            final CellStyle newStyle2 = newStyle;
            IntStream.rangeClosed(range.getFirstRow(), range.getLastRow()).parallel().forEach(rowIndex -> {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = sheet.createRow(rowIndex);
                }

                for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell(colIndex);
                    }

                    // セルに既存のスタイルを設定
                    cell.setCellStyle(newStyle2);
                }
            });

        }

        return newStyle;
    }

    // ラッパー関数
    public static CellStyle setCellBackgroundAndFontColor(Sheet sheet, CellRangeAddress range, Object bgColor, Object fontColor) {
        // baseStyle を null として呼び出すラッパー関数
        return setCellBackgroundAndFontColor(sheet, range, bgColor, fontColor, null);
    }

    /**
     * シートの範囲内のセルに背景色とフォント色を設定
     * 
     * @param sheet シート
     * @param range 範囲を表すCellRangeAddress
     * @param bgColor 背景色
     * @param fontColor フォント色
     * 
     *        使用例 ↓
     * 
     *        範囲を指定する
     *        CellRangeAddress range = new CellRangeAddress(2, 5, 3, 7); ※startRow, endRow, startCol, endCol
     *
     *        セルに背景色とフォント色を設定する
     *        setCellBackgroundAndFontColor(sheet, range, IndexedColors.LIGHT_GREEN, IndexedColors.WHITE);
     * 
     */
    // IndexedColorsを使用するメソッド
    public static void setCellBackgroundAndFontColor_void(Sheet sheet, CellRangeAddress range, Object bgColor, Object fontColor) {
        Workbook workbook = sheet.getWorkbook();

        // スタイルのキャッシュを作成
        Map<CellStyle, CellStyle> styleCache = new HashMap<>();

        for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }

                CellStyle existingStyle = cell.getCellStyle();
                CellStyle newStyle;

                // 既存のスタイルがキャッシュにあるか確認
                if (existingStyle != null && styleCache.containsKey(existingStyle)) {
                    newStyle = styleCache.get(existingStyle);
                } else {
                    newStyle = workbook.createCellStyle();
                    if (existingStyle != null) {
                        newStyle.cloneStyleFrom(existingStyle);
                    }

                    // 背景色を設定
                    if (bgColor instanceof IndexedColors) {
                        newStyle.setFillForegroundColor(((IndexedColors) bgColor).getIndex());
                    } else if (bgColor instanceof XSSFColor) {
                        newStyle.setFillForegroundColor((XSSFColor) bgColor);
                    }
                    newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                    // フォント色を設定
                    XSSFFont newFont = (XSSFFont) workbook.createFont();
                    Font existingFont = existingStyle != null ? workbook.getFontAt(existingStyle.getFontIndex()) : null;
                    if (existingFont != null) {
                        newFont.setFontName(existingFont.getFontName());
                        newFont.setFontHeightInPoints(existingFont.getFontHeightInPoints());
                        newFont.setBold(existingFont.getBold());
                        newFont.setItalic(existingFont.getItalic());
                        newFont.setStrikeout(existingFont.getStrikeout());
                        newFont.setTypeOffset(existingFont.getTypeOffset());
                        newFont.setUnderline(existingFont.getUnderline());
                    }

                    if (fontColor instanceof IndexedColors) {
                        newFont.setColor(((IndexedColors) fontColor).getIndex());
                    } else if (fontColor instanceof XSSFColor) {
                        throw new UnsupportedOperationException("XSSFColorの設定はバグがあり上手く機能しないので利用不可");
                    }

                    newStyle.setFont(newFont);
                    styleCache.put(existingStyle, newStyle);
                }

                // セルに新しいスタイルを設定
                cell.setCellStyle(newStyle);
            }
        }
    }

    /**
     * 指定された範囲のセルに、格子状の罫線を引く
     *
     * @param sheet シート
     * @param startRow 開始行インデックス
     * @param startColumn 開始列インデックス
     * @param endRow 終了行インデックス
     * @param endColumn 終了列インデックス
     */
    public static void setGridLines(Sheet sheet, int startRow, int startColumn, int endRow, int endColumn) {
        // 外側の罫線を設定
        CellRangeAddress region = new CellRangeAddress(startRow, endRow, startColumn, endColumn);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

        // ワークブックを取得
        Workbook workbook = sheet.getWorkbook();

        // スタイルのキャッシュを作成
        Map<CellStyle, CellStyle> styleCache = new HashMap<>();

        // 内側の罫線を設定
        for (int row = startRow; row <= endRow; row++) {
            Row sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            for (int col = startColumn; col <= endColumn; col++) {
                Cell cell = sheetRow.getCell(col);
                if (cell == null) {
                    cell = sheetRow.createCell(col);
                }
                CellStyle existingStyle = cell.getCellStyle();

                if (existingStyle != null) {
                    // 既存のスタイルがキャッシュにあるか確認
                    CellStyle cachedStyle = styleCache.get(existingStyle);
                    if (cachedStyle == null) {
                        // キャッシュにない場合、新しいスタイルを作成してキャッシュに追加
                        CellStyle newStyle = workbook.createCellStyle();
                        newStyle.cloneStyleFrom(existingStyle);
                        newStyle.setBorderBottom(BorderStyle.THIN);
                        newStyle.setBorderTop(BorderStyle.THIN);
                        newStyle.setBorderLeft(BorderStyle.THIN);
                        newStyle.setBorderRight(BorderStyle.THIN);
                        styleCache.put(existingStyle, newStyle);
                        cachedStyle = newStyle;
                    }
                    // セルにキャッシュされたスタイルを設定
                    cell.setCellStyle(cachedStyle);
                } else {
                    // 既存のスタイルがない場合、新しいスタイルを作成して適用
                    CellStyle newStyle = workbook.createCellStyle();
                    newStyle.setBorderBottom(BorderStyle.THIN);
                    newStyle.setBorderTop(BorderStyle.THIN);
                    newStyle.setBorderLeft(BorderStyle.THIN);
                    newStyle.setBorderRight(BorderStyle.THIN);
                    cell.setCellStyle(newStyle);
                }
            }
        }
    }

    // レンジ指定用のラッパー
    public static void setGridLines(Sheet sheet, CellRangeAddress range) {
        setGridLines(sheet, range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn());
    }

    // 指定列の末尾行を返す
    public static int getLastRowNumInColumn(Sheet sheet, int column) {
        int lastRowNum = -1;
        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(column);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    lastRowNum = rowIndex;
                }
            }
        }
        return lastRowNum;
    }

    /**
     * 指定された範囲のセルに、格子状の罫線を引く
     *
     * @param sheet シート
     * @param startRow 開始行インデックス
     * @param startColumn 開始列インデックス
     * @param endRow 終了行インデックス
     * @param endColumn 終了列インデックス
     */
    public static CellStyle setGridLines(Sheet sheet, int startRow, int startColumn, int endRow, int endColumn, CellStyle baseStyle) {
        CellStyle newStyle = baseStyle;

        if (newStyle == null) {
            // 外側の罫線を設定
            CellRangeAddress region = new CellRangeAddress(startRow, endRow, startColumn, endColumn);
            RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);

            // ワークブックを取得
            Workbook workbook = sheet.getWorkbook();
            Map<CellStyle, CellStyle> styleCache = new HashMap<>();

            // 内側の罫線を設定
            for (int row = startRow; row <= endRow; row++) {
                Row sheetRow = sheet.getRow(row);
                if (sheetRow == null) {
                    sheetRow = sheet.createRow(row);
                }
                for (int col = startColumn; col <= endColumn; col++) {
                    Cell cell = sheetRow.getCell(col);
                    if (cell == null) {
                        cell = sheetRow.createCell(col);
                    }
                    CellStyle existingStyle = cell.getCellStyle();

                    if (existingStyle != null) {
                        // 既存のスタイルがキャッシュにあるか確認
                        CellStyle cachedStyle = styleCache.get(existingStyle);
                        if (cachedStyle == null) {
                            // 既存のスタイルに今回設定するスタイルが既に適用されているかチェック
                            boolean bordersMatch = existingStyle.getBorderBottom() == BorderStyle.THIN &&
                                    existingStyle.getBorderTop() == BorderStyle.THIN &&
                                    existingStyle.getBorderLeft() == BorderStyle.THIN &&
                                    existingStyle.getBorderRight() == BorderStyle.THIN;

                            if (bordersMatch) {
                                // 既存のスタイルが今回のスタイルと一致する場合、そのまま流用
                                cachedStyle = existingStyle;
                            } else {
                                // 一致しない場合、新しいスタイルを作成してキャッシュに追加
                                newStyle = workbook.createCellStyle();
                                newStyle.cloneStyleFrom(existingStyle);
                                newStyle.setBorderBottom(BorderStyle.THIN);
                                newStyle.setBorderTop(BorderStyle.THIN);
                                newStyle.setBorderLeft(BorderStyle.THIN);
                                newStyle.setBorderRight(BorderStyle.THIN);
                                styleCache.put(existingStyle, newStyle);
                                cachedStyle = newStyle;
                            }
                        }
                        // セルにキャッシュされたスタイルを設定
                        cell.setCellStyle(cachedStyle);
                    } else {
                        // 既存のスタイルがない場合、新しいスタイルを作成して適用
                        newStyle = workbook.createCellStyle();
                        newStyle.setBorderBottom(BorderStyle.THIN);
                        newStyle.setBorderTop(BorderStyle.THIN);
                        newStyle.setBorderLeft(BorderStyle.THIN);
                        newStyle.setBorderRight(BorderStyle.THIN);
                        cell.setCellStyle(newStyle);
                    }
                }
            }
        } else {
            // 基本スタイルが渡されていた場合、そのまま使用する
            for (int row = startRow; row <= endRow; row++) {
                Row sheetRow = sheet.getRow(row);
                if (sheetRow == null) {
                    sheetRow = sheet.createRow(row);
                }
                for (int col = startColumn; col <= endColumn; col++) {
                    Cell cell = sheetRow.getCell(col);
                    if (cell == null) {
                        cell = sheetRow.createCell(col);
                    }

                    // セルに既存のスタイルを設定
                    cell.setCellStyle(newStyle);
                }
            }

//	    	final CellStyle newStyle2 = newStyle;
//	    	IntStream.rangeClosed(startRow, endRow).parallel().forEach(row -> {
//	    	    Row sheetRow = sheet.getRow(row);
//	    	    if (sheetRow == null) {
//	    	        sheetRow = sheet.createRow(row);
//	    	    }
//	    	    for (int col = startColumn; col <= endColumn; col++) {
//	    	        Cell cell = sheetRow.getCell(col);
//	    	        if (cell == null) {
//	    	            cell = sheetRow.createCell(col);
//	    	        }
//	    	        cell.setCellStyle(newStyle2);
//	    	    }
//	    	});	    	
        }

        return newStyle;
    }

    // baseStyle が無い場合のラッパー関数
    public static void setGridLines_void(Sheet sheet, int startRow, int startColumn, int endRow, int endColumn) {
        setGridLines(sheet, startRow, startColumn, endRow, endColumn, null);
    }

    // レンジ指定用のラッパー
    public static void setGridLines_void(Sheet sheet, CellRangeAddress range) {
        setGridLines(sheet, range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn());
    }

    public static CellStyle setGridLines(Sheet sheet, CellRangeAddress range, CellStyle baseStyle) {
        return setGridLines(sheet, range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn(), baseStyle);
    }

    // スタイルを適用する
    public static void setCellStyle(Sheet sheet, int startRow, int startColumn, int endRow, int endColumn, CellStyle style) {
        for (int row = startRow; row <= endRow; row++) {
            Row sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            for (int col = startColumn; col <= endColumn; col++) {
                Cell cell = sheetRow.getCell(col);
                if (cell == null) {
                    cell = sheetRow.createCell(col);
                }

                // セルに既存のスタイルを設定
                cell.setCellStyle(style);
            }
        }
    }

    public static void setCellStyle(Sheet sheet, CellRangeAddress range, CellStyle style) {
        setCellStyle(sheet, range.getFirstRow(), range.getFirstColumn(), range.getLastRow(), range.getLastColumn(), style);
    }

    /**
     * シート内の最後の列のインデックスを取得
     * 
     * @param sheet シート
     * @return 最後の列のインデックス
     */
    public static int getLastColumnIndex(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        int lastColNum = 0;
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int lastCellNum = row.getLastCellNum();
                if (lastCellNum > lastColNum) {
                    lastColNum = lastCellNum;
                }
            }
        }
        return lastColNum - 1; // 0から始まるインデックスに変換する
    }

    public static void printCellStyleCount(Sheet sheet) {
        // セットを使ってユニークなスタイルをカウント
        Set<CellStyle> uniqueStyles = new HashSet<>();

        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int colIndex = row.getFirstCellNum(); colIndex < row.getLastCellNum(); colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null) {
                        CellStyle style = cell.getCellStyle();
                        if (style != null) {
                            uniqueStyles.add(style);
                        }
                    }
                }
            }
        }

        // コンソールにスタイルの数を出力
        System.out.println("Sheet '" + sheet.getSheetName() + "' contains " + uniqueStyles.size() + " unique cell styles.");
    }

    // セルに取り消し線を付ける
    public static void addStrikethroughToCell(Workbook workbook, Cell cell) {
        if (cell == null) return;
        CellStyle existingStyle = cell.getCellStyle();
        Font existingFont = workbook.getFontAt(existingStyle.getFontIndex());

        // 新しいフォントを作成して既存のフォントのプロパティをコピー
        Font newFont = workbook.createFont();
        newFont.setFontName(existingFont.getFontName());
        newFont.setFontHeight(existingFont.getFontHeight());
        newFont.setBold(existingFont.getBold());
        newFont.setItalic(existingFont.getItalic());
        newFont.setStrikeout(true); // 取り消し線を追加
        newFont.setColor(existingFont.getColor());
        newFont.setUnderline(existingFont.getUnderline());
        newFont.setTypeOffset(existingFont.getTypeOffset());

        // 新しいスタイルを作成して既存のスタイルのプロパティをコピー
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(existingStyle);
        newStyle.setFont(newFont);

        // 新しいスタイルをセルに適用
        cell.setCellStyle(newStyle);
    }

}
