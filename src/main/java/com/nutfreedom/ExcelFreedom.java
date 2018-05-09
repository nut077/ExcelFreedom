package com.nutfreedom;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Vector;

import javax.imageio.ImageIO;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.jsp.JspWriter;

import jxl.CellView;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

public class ExcelFreedom {
    private LinkedHashMap<String, LinkedHashMap<String, String>> detailFormat = new LinkedHashMap<>();
    private LinkedHashMap<String, LinkedHashMap<String, String>> detailFont = new LinkedHashMap<>();
    private LinkedHashMap<String, LinkedHashMap<String, String>> detailImage = new LinkedHashMap<>();
    private LinkedHashMap<String, Integer> widthCell = new LinkedHashMap<>();
    private LinkedHashMap<String, Integer> heightCell = new LinkedHashMap<>();
    private Vector<String> span = new Vector<>();
    private String table;
    private String locationFile;
    private WritableWorkbook workbook;
    private ServletOutputStream locationFileServlet;
    private int defaultFontSize = 10;
    private String defaultFormat = "border-center";
    private String defaultFontColor = "black";
    private String defaultFontName = "arial";
    private int defaultHeight = 0;
    private String spacebar = "false";
    private boolean statusFile = true;
    private boolean statusAutoSize = true;
    private static final double CELL_DEFAULT_HEIGHT = 17;
    private static final double CELL_DEFAULT_WIDTH = 64;
    private ParseNumberFreedom parse = new ParseNumberFreedom();
    private SplitFreedom split = new SplitFreedom();
    private FormatExcelFreedom format = new FormatExcelFreedom();
    private CheckFreedom check = new CheckFreedom();
    private SubstringFreedom sub = new SubstringFreedom();
    private HttpServletResponse response = null;
    private String filename = "";
    private JspWriter out = null;

    public ExcelFreedom() {
    }

    public ExcelFreedom(String table, String locationFile) {
        this.table = table;
        this.locationFile = locationFile;
    }

    public ExcelFreedom(String table, ServletOutputStream locationFileServlet) {
        this.table = table;
        this.locationFileServlet = locationFileServlet;
        this.statusFile = false;
    }

    public ExcelFreedom(String table, ServletOutputStream locationFileServlet, HttpServletResponse response, JspWriter out, String filename) {
        this.table = table;
        this.locationFileServlet = locationFileServlet;
        this.response = response;
        this.out = out;
        this.filename = filename;
        this.statusFile = false;
    }

    public void setTable(String table) {
        this.table = table;
    }

    public String getTable() {
        return this.table;
    }

    public String getLocationFile() {
        return locationFile;
    }

    public void setLocationFile(String locationFile) {
        this.locationFile = locationFile;
    }

    public int getDefaultFontSize() {
        return this.defaultFontSize;
    }

    public void setDefaultFontSize(int defaultFontSize) {
        this.defaultFontSize = defaultFontSize;
    }

    public String getDefaultFontColor() {
        return defaultFontColor;
    }

    public void setDefaultFontColor(String defaultFontColor) {
        this.defaultFontColor = defaultFontColor;
    }

    public String getDefaultFontName() {
        return defaultFontName;
    }

    public void setDefaultFontName(String defaultFontName) {
        this.defaultFontName = defaultFontName;
    }

    public int getDefaultHeight() {
        return defaultHeight;
    }

    public void setDefaultHeight(int defaultHeight) {
        this.defaultHeight = defaultHeight;
    }

    private void setWorkbook() {
        try {
            this.workbook = Workbook.createWorkbook(new File(this.locationFile));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void setWorkbookServlet() {
        try {
            this.workbook = Workbook.createWorkbook(this.locationFileServlet);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public String getDefaultFormat() {
        return defaultFormat;
    }

    public void setDefaultFormat(String defaultFormat) {
        this.defaultFormat = defaultFormat;
    }

    public void setAutoSize(boolean statusAutoSize) {
        this.statusAutoSize = statusAutoSize;
    }

    public String getSpacebar() {
        return spacebar;
    }

    public void setSpacebar(String spacebar) {
        this.spacebar = spacebar;
    }

    private Integer chkCellDuplicate(Vector<String> used, String colRow, int plusCol) {
        int tmpCol = 0;
        boolean isDuplicate = false;
        int col = parse.parseInt(split.splitByIndex(colRow, "#", 0));
        while (used.contains(colRow)) {
            String[] sp = split.split(colRow, "#");
            tmpCol = parse.parseInt(sp[0]) + 1;
            colRow = tmpCol + "#" + sp[1];
            plusCol++;
            isDuplicate = true;
        }
        if (isDuplicate) {
            return tmpCol;
        } else {
            return col;
        }
    }

    private LinkedHashMap<String, String> chkCellDuplicate(Vector<String> used, LinkedHashMap<String, String> hm, String colRow, int plusCol) {
        hm.put("plusCol", String.valueOf(plusCol));
        while (used.contains(colRow)) {
            String[] sp = split.split(colRow, "#");
            String tmpCol = String.valueOf(parse.parseInt(sp[0]) + 1);
            colRow = tmpCol + "#" + sp[1];
            plusCol++;
            hm.put("col", tmpCol);
            hm.put("plusCol", String.valueOf(plusCol));
        }
        return hm;
    }

    public void write() {
        if (statusFile) {
            this.setWorkbook();
        } else {
            this.setWorkbookServlet();
        }
        writingToExcel(this.workbook);
    }

    private void writingToExcel(WritableWorkbook workbook) {
        try {
            createContent(workbook);
        } finally {
            try {
                workbook.write();
                workbook.close();
            } catch (IOException | WriteException e) {
                e.printStackTrace();
            }
        }
    }

    private void createContent(WritableWorkbook workbook) {
        int excelSheetIndex = 0;
        WritableSheet sheet;
        String[] table = split.split(this.table, "<table>");
        for (int tableRow = 1; tableRow < table.length; tableRow++) {
            clearDetail();
            this.setDetail(table[tableRow]);
            String sheetName = "sheet " + String.valueOf(excelSheetIndex + 1);
            if (check.isHaveData(table[tableRow], "<sheet>")) {
                sheetName = check.setValueNotBlank(getValueConditionString(table[tableRow], "<sheet>", "</sheet>"), sheetName);
            }
            sheet = workbook.createSheet(sheetName, excelSheetIndex++);
            process(sheet);
            if (this.statusAutoSize) {
                setAutoSize(sheet);
            }
            setWidthCell(sheet);
            setHeightCell(sheet);

            if (response != null) {
                response.setContentType("application/vnd.ms-excel");
                response.setHeader("Pragma", "public");
                response.setHeader("Cache-Control", "max-age=0");
                response.setHeader("Content-Disposition", "attachment; filename=" + filename + ".xls");
                try {
                    out.clear();
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
        }
    }

    private void clearDetail() {
        this.detailFont.clear();
        this.detailFormat.clear();
        this.detailImage.clear();
        this.widthCell.clear();
        this.heightCell.clear();
        this.span.clear();
    }

    private void setDetail(String table) {
        Vector<String> used = new Vector<>();
        String[] tr = split.split(table, "<tr>");
        for (int trRow = 1; trRow < tr.length; trRow++) {
            String[] td = split.split(tr[trRow], "<td>");
            int plusCol = 0;
            for (int tdCol = 1; tdCol < td.length; tdCol++) {
                int col = tdCol - 1 + plusCol;
                int row = trRow - 1;
                boolean isColspan = false;
                int tmpPlusCol = 0;
                String data = sub.substring(td[tdCol], 0, td[tdCol].indexOf("</td>"));
                if (data.equals("") && (!td[tdCol].equals("</td>") && !td[tdCol].equals("</td></tr>") && !td[tdCol].equals("</td></tr></table>")) ) {
                    printError("expect </td> but this code => " + td[tdCol]);
                }
                if (chkCondition(data, "<colspan>") && chkCondition(data, "<rowspan>")) {
                    int colspan = getValueConditionInt(data, "<colspan>", "</colspan>") - 1;
                    plusCol += colspan;
                    int rowspan = getValueConditionInt(data, "<rowspan>", "</rowspan>") - 1;
                    this.span.addElement(col + "#" + row + "#" + String.valueOf(col + colspan) + "#" + String.valueOf(row + rowspan));

                    boolean first = true;
                    for (int i = col; i <= col + colspan; i++) {
                        int line = 0;
                        for (int j = row; j <= row + rowspan; j++) {
                            if (first) {
                                first = false;
                            } else {
                                used.addElement(i + "#" + String.valueOf(row + line));
                            }
                            line++;
                        }
                    }

                } else if (chkCondition(data, "<colspan>")) {
                    int colspan = getValueConditionInt(data, "<colspan>", "</colspan>") - 1;
                    plusCol += colspan;
                    col = chkCellDuplicate(used, col + "#" + row, plusCol);
                    this.span.addElement(col + "#" + row + "#" + String.valueOf(col + colspan) + "#" + row);
                    for (int i = col; i < col + colspan; i++) {
                        used.addElement(i + "#" + row);
                    }
                    tmpPlusCol = colspan;
                    isColspan = true;
                } else if (chkCondition(data, "<rowspan>")) {
                    int rowspan = getValueConditionInt(data, "<rowspan>", "</rowspan>") - 1;
                    col = chkCellDuplicate(used, col + "#" + row, plusCol);
                    this.span.addElement(col + "#" + row + "#" + col + "#" + String.valueOf(row + rowspan));

                    int tmp = 1;
                    for (int i = row; i < row + rowspan; i++) {
                        used.addElement(col + "#" + String.valueOf(row + tmp));
                        tmp++;
                    }
                }

                String format = defaultFormat;
                if (chkCondition(data, "<format>")) {
                    format = getValueConditionString(data, "<format>", "</format>");
                }

                String type = "string";
                if (chkCondition(data, "<type>")) {
                    type = getValueConditionString(data, "<type>", "</type>");
                }

                boolean statusFont = false;
                String fontname = defaultFontName;
                if (chkCondition(data, "<font-name>")) {
                    fontname = getValueConditionString(data, "<font-name>", "</font-name>");
                    statusFont = true;
                }
                String fontsize = String.valueOf(defaultFontSize);
                if (chkCondition(data, "<font-size>")) {
                    fontsize = getValueConditionString(data, "<font-size>", "</font-size>");
                    statusFont = true;
                }
                String bold = "false";
                if (chkCondition(data, "<b>")) {
                    bold = getValueConditionString(data, "<b>", "</b>");
                    statusFont = true;
                }
                String italic = "false";
                if (chkCondition(data, "<i>")) {
                    italic = getValueConditionString(data, "<i>", "</i>");
                    statusFont = true;
                }
                String underline = "no_underline";
                if (chkCondition(data, "<u>")) {
                    underline = getValueConditionString(data, "<u>", "</u>");
                    statusFont = true;
                }
                String color = defaultFontColor;
                if (chkCondition(data, "<color>")) {
                    color = getValueConditionString(data, "<color>", "</color>");
                    statusFont = true;
                }
                String background = "";
                if (chkCondition(data, "<background>")) {
                    background = getValueConditionString(data, "<background>", "</background>");
                    statusFont = true;
                }

                String wrap = "false";
                if (chkCondition(data, "<wrap>")) {
                    wrap = getValueConditionString(data, "<wrap>", "</wrap>");
                    statusFont = true;
                }

                String width = "";
                if (chkCondition(data, "<width>")) {
                    width = getValueConditionString(data, "<width>", "</width>");
                }

                int height = defaultHeight;
                if (chkCondition(data, "<height>")) {
                    height = getValueConditionInt(data, "<height>", "</height>");
                }

                String formula = "false";
                if (chkCondition(data, "<formula>")) {
                    formula = getValueConditionString(data, "<formula>", "</formula>");
                }

                // SUM(#1@1#:#2@1#) == SUM(A1:B1)
                String formulaNo = "false";
                if (chkCondition(data, "<formula-no>")) {
                    formulaNo = getValueConditionString(data, "<formula-no>", "</formula-no>");
                }

                if (chkCondition(data, "<spacebar>")) {
                    spacebar = getValueConditionString(data, "<spacebar>", "</spacebar>");
                }

                boolean statusImage = false;
                String imageWidth = "0";
                String imageHeight = "0";
                if (chkCondition(data, "<image>")) {
                    if (chkCondition(data, "<image-width>")) {
                        imageWidth = getValueConditionString(data, "<image-width>", "</image-width>");
                    }
                    if (chkCondition(data, "<image-height>")) {
                        imageHeight = getValueConditionString(data, "<image-height>", "</image-height>");
                    }
                    data = getValueConditionString(data, "<image>", "</image>");
                    statusImage = true;
                }

                if (chkCondition(data, "<") && !statusImage) {
                    String tmp = sub.substring(data, data.lastIndexOf("</") + 2, data.length());
                    data = sub.substring(tmp, tmp.indexOf(">") + 1, tmp.length());
                }

                LinkedHashMap<String, String> hm = new LinkedHashMap<>();
                hm.put("col", String.valueOf(col));
                hm.put("row", String.valueOf(row));
                hm.put("data", data);
                hm.put("type", type);
                hm.put("format", format);
                hm.put("formula", formula);
                hm.put("formula-no", formulaNo);
                hm.put("spacebar", spacebar);

                String colRow = col + "#" + row;
                if (isColspan) {
                    plusCol -= tmpPlusCol;
                }
                hm = chkCellDuplicate(used, hm, colRow, plusCol);
                colRow = hm.get("col") + "#" + hm.get("row");
                if (isColspan) {
                    hm.put("col", String.valueOf(col));
                }
                if (statusFont) {
                    hm.put("font-name", fontname);
                    hm.put("font-size", fontsize);
                    hm.put("bold", bold);
                    hm.put("italic", italic);
                    hm.put("underline", underline);
                    hm.put("color", color);
                    hm.put("background", background);
                    hm.put("wrap", wrap);
                    this.detailFont.put(colRow, hm);
                } else if (statusImage) {
                    hm.put("image", data);
                    hm.put("image-width", imageWidth);
                    hm.put("image-height", imageHeight);
                    this.detailImage.put(colRow, hm);
                } else {
                    this.detailFormat.put(colRow, hm);
                }

                if (!width.equals("")) {
                    this.widthCell.put(hm.get("col"), parse.parseInt(width));
                }

                if (check.isMoreThanZero(height)) {
                    this.heightCell.put(hm.get("row"), height);
                }

                plusCol = parse.parseInt(hm.get("plusCol"));
                used.addElement(colRow);
            }
        }
    }

    private void process(WritableSheet sheet) {

        this.format.setDefaultFontSize(this.defaultFontSize);
        this.format.setDefaultFontColor(this.defaultFontColor);
        this.format.setDefaultFontName(this.defaultFontName);
        this.format.process();

        for (String key : this.detailFormat.keySet()) {
            HashMap<String, String> hm = this.detailFormat.get(key);
            String data = hm.get("data");
            int col = parse.parseInt(hm.get("col"));
            int row = parse.parseInt(hm.get("row"));
            String format = hm.get("format");
            String type = hm.get("type");
            String formula = hm.get("formula");
            String formulaNo = hm.get("formula-no");
            switch (type) {
                case "number":
                    if (formula.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, data, format);
                    } else if (formulaNo.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, getNewColumn(data), format);
                    } else {
                        setDataCellFormatNumber(sheet, col, row, parse.parseDouble(data), format);
                    }
                    break;
                case "number-money":
                    if (formula.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, data, "number-" + format);
                    } else if (formulaNo.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, getNewColumn(data), "number-" + format);
                    } else {
                        setDataCellFormatNumber(sheet, col, row, parse.parseDouble(data), "number-" + format);
                    }
                    break;
                case "number-money-float":
                    if (formula.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, data, "number-float-" + format);
                    } else if (formulaNo.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, getNewColumn(data), "number-float-" + format);
                    } else {
                        setDataCellFormatNumber(sheet, col, row, parse.parseDouble(data), "number-float-" + format);
                    }
                    break;
                case "number-money-float-one":
                    if (formula.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, data, "number-float-one-" + format);
                    } else if (formulaNo.equals("true")) {
                        setDataCellFormatNumberFormula(sheet, col, row, getNewColumn(data), "number-float-one-" + format);
                    } else {
                        setDataCellFormatNumber(sheet, col, row, parse.parseDouble(data), "number-float-one-" + format);
                    }
                    break;
                default:
                    if (hm.get("spacebar").equals("true")) {
                        data = " " + data + " ";
                    }
                    setDataCellFormatString(sheet, col, row, data, format, formula, formulaNo);
                    break;
            }
        }

        for (String key : this.detailFont.keySet()) {
            HashMap<String, String> hm = this.detailFont.get(key);
            String type = hm.get("type");
            String formula = hm.get("formula");
            String formulaNo = hm.get("formula-no");
            switch (type) {
                case "number":
                    if (formula.equals("true")) {
                        setDataCellFontNumberFormula(sheet, hm, "");
                    } else if (formulaNo.equals("true")) {
                        setDataCellFontNumberFormulaNo(sheet, hm, "");
                    } else {
                        setDataCellFontNumber(sheet, hm, "");
                    }
                    break;
                case "number-money":
                    if (formula.equals("true")) {
                        setDataCellFontNumberFormula(sheet, hm, "number-");
                    } else if (formulaNo.equals("true")) {
                        setDataCellFontNumberFormulaNo(sheet, hm, "number-");
                    } else {
                        setDataCellFontNumber(sheet, hm, "number-");
                    }
                    break;
                case "number-money-float":
                    if (formula.equals("true")) {
                        setDataCellFontNumberFormula(sheet, hm, "number-float-");
                    } else if (formulaNo.equals("true")) {
                        setDataCellFontNumberFormulaNo(sheet, hm, "number-float-");
                    } else {
                        setDataCellFontNumber(sheet, hm, "number-float-");
                    }
                    break;
                case "number-money-float-one":
                    if (formula.equals("true")) {
                        setDataCellFontNumberFormula(sheet, hm, "number-float-one-");
                    } else if (formulaNo.equals("true")) {
                        setDataCellFontNumberFormulaNo(sheet, hm, "number-float-one-");
                    } else {
                        setDataCellFontNumber(sheet, hm, "number-float-one-");
                    }
                    break;
                default:
                    setDataCellFontString(sheet, hm);
                    break;
            }
        }

        for (String key : this.detailImage.keySet()) {
            HashMap<String, String> hm = this.detailImage.get(key);
            addImage(sheet, parse.parseInt(hm.get("col")), parse.parseInt(hm.get("row")), hm.get("image"), parse.parseDouble(hm.get("image-width")), parse.parseDouble(hm.get("image-height")));
        }

        for (String value : this.span) {
            setMergeCell(sheet, value);
        }
    }

    private String getNewColumn(String val) {
        StringBuilder result = new StringBuilder();
        String str;
        int col;
        int row;
        String[] sp = split.split(val, "#");
        for (String aSp : sp) {
            if (check.isHaveData(aSp, "@")) {
                String[] colRow = split.split(aSp, "@");
                col = parse.parseInt(colRow[0]);
                row = parse.parseInt(colRow[1]);
                int first = (int) Math.ceil(col / 26.0);
                if (first > 1) {
                    str = (char) (63 + first) + "";
                    int remain = col - (26 * (first - 1));
                    str += (char) (remain + 64) + "";
                } else {
                    str = (char) (col + 64) + "";
                }
                result.append(str).append(row);
            } else {
                result.append(aSp);
            }
        }
        return result.toString();
    }

    private void addImage(WritableSheet sheet, int col, int row, String location, double width, double height) {
        try {
            File imageFile = new File(location);
            BufferedImage input;
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            input = ImageIO.read(imageFile);
            ImageIO.write(input, "PNG", baos);

            double imageWidth = input.getWidth() / CELL_DEFAULT_WIDTH;
            double imageHeight = input.getHeight() / CELL_DEFAULT_HEIGHT;
            if (width != 0) {
                imageWidth = width;
            }
            if (height != 0) {
                imageHeight = height;
            }
            sheet.addImage(new WritableImage(col, row, imageWidth, imageHeight, baos.toByteArray()));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFormatNumber(WritableSheet sheet, int col, int row, double data, String ft) {
        try {
            sheet.addCell(new Number(col, row, data, this.format.getFormat(ft)));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFormatNumberFormula(WritableSheet sheet, int col, int row, String data, String ft) {
        try {
            sheet.addCell(new Formula(col, row, data, this.format.getFormat(ft)));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFormatString(WritableSheet sheet, int col, int row, String data, String ft, String formula, String formulaNo) {
        try {
            if (formula.equals("false")) {
                sheet.addCell(new Label(col, row, data, this.format.getFormat(ft)));
            } else if (formulaNo.equals("true")) {
                sheet.addCell(new Formula(col, row, getNewColumn(data), this.format.getFormat(ft)));
            } else {
                sheet.addCell(new Formula(col, row, data, this.format.getFormat(ft)));
            }
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFontString(WritableSheet sheet, HashMap<String, String> hm) {
        try {
            WritableCellFormat cell = setCell(hm, this.format.getFormat(hm.get("format")));
            String data = hm.get("data");
            int col = parse.parseInt(hm.get("col"));
            int row = parse.parseInt(hm.get("row"));
            if (hm.get("formula").equals("false")) {
                sheet.addCell(new Label(col, row, data, cell));
            } else if (hm.get("formula-no").equals("true")) {
                sheet.addCell(new Formula(col, row, getNewColumn(data), cell));
            } else {
                if (hm.get("spacebar").equals("true")) {
                    data = " " + data + " ";
                }
                sheet.addCell(new Formula(col, row, data, cell));
            }
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFontNumber(WritableSheet sheet, HashMap<String, String> hm, String formatNumberMoney) {
        try {
            WritableCellFormat cell = setCell(hm, this.format.getFormat(formatNumberMoney + hm.get("format")));
            sheet.addCell(new Number(parse.parseInt(hm.get("col")), parse.parseInt(hm.get("row")), parse.parseDouble(hm.get("data")), cell));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFontNumberFormula(WritableSheet sheet, HashMap<String, String> hm, String formatNumberMoney) {
        try {
            WritableCellFormat cell = setCell(hm, this.format.getFormat(formatNumberMoney + hm.get("format")));
            sheet.addCell(new Formula(parse.parseInt(hm.get("col")), parse.parseInt(hm.get("row")), hm.get("data"), cell));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setDataCellFontNumberFormulaNo(WritableSheet sheet, HashMap<String, String> hm, String formatNumberMoney) {
        try {
            WritableCellFormat cell = setCell(hm, this.format.getFormat(formatNumberMoney + hm.get("format")));
            sheet.addCell(new Formula(parse.parseInt(hm.get("col")), parse.parseInt(hm.get("row")), getNewColumn(hm.get("data")), cell));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private WritableCellFormat setCell(HashMap<String, String> hm, WritableCellFormat format) {
        WritableCellFormat cell = new WritableCellFormat(format);
        try {
            if (hm.get("bold").equals("true")) {
                cell.setFont(new WritableFont(this.format.getFontName(hm.get("font-name")), parse.parseInt(hm.get("font-size")), WritableFont.BOLD, hm.get("italic").equals("true"), this.format.getUnderline(hm.get("underline")), this.format.getColor(hm.get("color"))));
                if (!hm.get("background").equals("")) {
                    cell.setBackground(this.format.getColor(hm.get("background")));
                }
                setWrap(cell, hm.get("wrap"));
            } else {
                cell.setFont(new WritableFont(this.format.getFontName(hm.get("font-name")), parse.parseInt(hm.get("font-size")), WritableFont.NO_BOLD, hm.get("italic").equals("true"), this.format.getUnderline(hm.get("underline")), this.format.getColor(hm.get("color"))));
                if (!hm.get("background").equals("")) {
                    cell.setBackground(this.format.getColor(hm.get("background")));
                }
                setWrap(cell, hm.get("wrap"));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return cell;
    }

    private void setWrap(WritableCellFormat cell, String wrap) {
        try {
            cell.setWrap(wrap.equals("true"));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setMergeCell(WritableSheet sheet, String col) {
        try {
            String[] sp = split.split(col, "#");
            sheet.mergeCells(parse.parseInt(sp[0]), parse.parseInt(sp[1]), parse.parseInt(sp[2]), parse.parseInt(sp[3]));
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    private void setWidthCell(WritableSheet sheet) {
        for (String key : this.widthCell.keySet()) {
            int value = this.widthCell.get(key);
            int col = parse.parseInt(key);
            sheet.setColumnView(col, value);
        }
    }

    private void setHeightCell(WritableSheet sheet) {
        for (String key : this.heightCell.keySet()) {
            int value = this.heightCell.get(key);
            int row = parse.parseInt(key);
            try {
                sheet.setRowView(row, value);
            } catch (RowsExceededException e) {
                e.printStackTrace();
            }
        }
    }

    private void setAutoSize(WritableSheet sheet) {
        for (int i = 0; i < sheet.getColumns(); i++) {
            if (!this.widthCell.containsKey(String.valueOf(i))) {
                CellView cv = sheet.getColumnView(i);
                cv.setAutosize(true);
                sheet.setColumnView(i, cv);
            }
        }
    }

    private boolean chkCondition(String data, String con) {
        return check.isHaveData(data, con);
    }

    private int getValueConditionInt(String data, String start, String end) {
        String num = sub.substring(data, data.indexOf(start) + start.length(), data.indexOf(end));
        if (check.isBlank(num)) {
            printError("tag => " + start + " : " + end + "\n data => " + data);
            return 1;
        }
        return parse.parseInt(num);
    }

    private String getValueConditionString(String data, String start, String end) {
        String val = sub.substring(data, data.indexOf(start) + start.length(), data.indexOf(end));
        if (check.isBlank(val)) {
            printError("tag => " + start + " : " + end + "\n data => " + data);
        }
        return val;
    }

    private void printError(String msg) {
        System.out.println("ExcelFreedom error => " + msg);
    }

}