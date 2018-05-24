package com.nutfreedom.excel;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import jxl.biff.DisplayFormat;
import jxl.format.*;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;

import jxl.write.WritableFont.FontName;

class FormatExcelFreedom {
    private LinkedHashMap<String, WritableCellFormat> format = new LinkedHashMap<>();
    private LinkedHashMap<String, FontName> fontName = new LinkedHashMap<>();
    private LinkedHashMap<String, UnderlineStyle> underline = new LinkedHashMap<>();
    private LinkedHashMap<String, Colour> color = new LinkedHashMap<>();
    private WritableFont font = null;

    private int defaultFontSize;
    private String defaultFontColor;
    private String defaultFontName;

    FormatExcelFreedom() {

    }

    void process() {
        this.setFontName();
        this.setUnderline();
        this.setColor();
        this.setFont();
        this.setFormat();
    }

    private void setFontName() {
        this.fontName.put("arial", WritableFont.ARIAL);
        this.fontName.put("tahoma", WritableFont.TAHOMA);
        this.fontName.put("courier", WritableFont.COURIER);
        this.fontName.put("times", WritableFont.TIMES);
        this.fontName.put("thsarabun_spk", WritableFont.createFont("TH SarabunPSK"));
        this.fontName.put("thsarabun_new", WritableFont.createFont("TH Sarabun New"));
    }

    public HashMap<String, FontName> getFontName() {
        return fontName;
    }

    FontName getFontName(String fontName) {
        if (this.fontName.containsKey(fontName)) {
            return this.fontName.get(fontName);
        } else {
            printError(fontName, this.fontName);
            return this.fontName.get("arial");
        }
    }

    private void setUnderline() {
        this.underline.put("no_underline", UnderlineStyle.NO_UNDERLINE);
        this.underline.put("double", UnderlineStyle.DOUBLE);
        this.underline.put("double_accounting", UnderlineStyle.DOUBLE_ACCOUNTING);
        this.underline.put("single", UnderlineStyle.SINGLE);
    }

    UnderlineStyle getUnderline(String underline) {
        if (this.underline.containsKey(underline)) {
            return this.underline.get(underline);
        } else {
            printError(underline, this.underline);
            return this.underline.get("no_underline");
        }
    }

    private void setColor() {
        this.color.put("aqua", Colour.AQUA);
        this.color.put("black", Colour.AUTOMATIC);
        this.color.put("blue", Colour.BLUE);
        this.color.put("blue_grey", Colour.BLUE_GREY);
        this.color.put("bright_green", Colour.BRIGHT_GREEN);
        this.color.put("brown", Colour.BROWN);
        this.color.put("coral", Colour.CORAL);
        this.color.put("dark_blue", Colour.DARK_BLUE);
        this.color.put("dark_green", Colour.DARK_GREEN);
        this.color.put("dark_purple", Colour.DARK_PURPLE);
        this.color.put("dark_red", Colour.DARK_RED);
        this.color.put("dark_teal", Colour.DARK_TEAL);
        this.color.put("dark_yellow", Colour.DARK_YELLOW);
        this.color.put("gold", Colour.GOLD);
        this.color.put("gray_25", Colour.GRAY_25);
        this.color.put("gray_50", Colour.GRAY_50);
        this.color.put("gray_80", Colour.GRAY_80);
        this.color.put("grey_25_percent", Colour.GREY_25_PERCENT);
        this.color.put("grey_40_percent", Colour.GREY_40_PERCENT);
        this.color.put("grey_50_percent", Colour.GREY_50_PERCENT);
        this.color.put("grey_80_percent", Colour.GREY_80_PERCENT);
        this.color.put("green", Colour.GREEN);
        this.color.put("ice_blue", Colour.ICE_BLUE);
        this.color.put("indigo", Colour.INDIGO);
        this.color.put("ivory", Colour.IVORY);
        this.color.put("lavender", Colour.LAVENDER);
        this.color.put("light_blue", Colour.LIGHT_BLUE);
        this.color.put("light_green", Colour.LIGHT_GREEN);
        this.color.put("light_orange", Colour.LIGHT_ORANGE);
        this.color.put("light_turquoise", Colour.LIGHT_TURQUOISE);
        this.color.put("lime", Colour.LIME);
        this.color.put("ocean_blue", Colour.OCEAN_BLUE);
        this.color.put("ocean_green", Colour.OLIVE_GREEN);
        this.color.put("orange", Colour.ORANGE);
        this.color.put("pale_blue", Colour.PALE_BLUE);
        this.color.put("periwinkle", Colour.PERIWINKLE);
        this.color.put("pink", Colour.PINK);
        this.color.put("plum", Colour.PLUM);
        this.color.put("red", Colour.RED);
        this.color.put("rose", Colour.ROSE);
        this.color.put("sea_green", Colour.SEA_GREEN);
        this.color.put("sky_blue", Colour.SKY_BLUE);
        this.color.put("tan", Colour.TAN);
        this.color.put("teal", Colour.TEAL);
        this.color.put("turquoise", Colour.TURQUOISE);
        this.color.put("very_light_yellow", Colour.VERY_LIGHT_YELLOW);
        this.color.put("violet", Colour.VIOLET);
        this.color.put("white", Colour.WHITE);
        this.color.put("yellow", Colour.YELLOW);
    }

    Colour getColor(String color) {
        if (this.color.containsKey(color)) {
            return this.color.get(color);
        } else {
            printError(color, this.color);
            return this.color.get("black");
        }
    }

    private void setFormat() {
        this.format.put("left", setFormatCell(Alignment.LEFT));
        this.format.put("center", setFormatCell(Alignment.CENTRE));
        this.format.put("right", setFormatCell(Alignment.RIGHT));
        this.format.put("left-middle", setFormatCell(Alignment.LEFT, VerticalAlignment.CENTRE));
        this.format.put("center-middle", setFormatCell(Alignment.CENTRE, VerticalAlignment.CENTRE));
        this.format.put("right-middle", setFormatCell(Alignment.RIGHT, VerticalAlignment.CENTRE));
        this.format.put("left-top", setFormatCell(Alignment.LEFT, VerticalAlignment.TOP));
        this.format.put("center-top", setFormatCell(Alignment.CENTRE, VerticalAlignment.TOP));
        this.format.put("right-top", setFormatCell(Alignment.RIGHT, VerticalAlignment.TOP));
        this.format.put("orientation", setFormatCellOrientation());
        this.format.put("orientation-middle", setFormatCellOrientation(VerticalAlignment.CENTRE));
        this.format.put("orientation-top", setFormatCellOrientation(VerticalAlignment.TOP));

        this.format.put("border-orientation", setFormatCellOrientationBorder(BorderLineStyle.THIN));
        this.format.put("border-dashed-orientation", setFormatCellOrientationBorder(BorderLineStyle.DASHED));
        this.format.put("border-orientation-middle", setFormatCellOrientationBorder(VerticalAlignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-dashed-orientation-middle", setFormatCellOrientationBorder(VerticalAlignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-orientation-top", setFormatCellOrientationBorder(VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-dashed-orientation-top", setFormatCellOrientationBorder(VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-tb-orientation", setFormatCellOrientationBorderTopBottom(BorderLineStyle.THIN));
        this.format.put("border-dashed-tb-orientation", setFormatCellOrientationBorderTopBottom(BorderLineStyle.DASHED));
        this.format.put("border-tb-orientation-middle", setFormatCellOrientationBorderTopBottom(VerticalAlignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-dashed-tb-orientation-middle", setFormatCellOrientationBorderTopBottom(VerticalAlignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-tb-orientation-top", setFormatCellOrientationBorderTopBottom(VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-dashed-tb-orientation-top", setFormatCellOrientationBorderTopBottom(VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-tb-double-orientation", setFormatCellOrientationBorderTopBottomDouble());
        this.format.put("border-tb-double-orientation-middle", setFormatCellOrientationBorderTopBottomDouble(VerticalAlignment.CENTRE));
        this.format.put("border-tb-double-orientation-top", setFormatCellOrientationBorderTopBottomDouble(VerticalAlignment.TOP));
        this.format.put("border-tb-single-double-orientation", setFormatCellOrientationBorderTopBottomSingleDouble());
        this.format.put("border-tb-single-double-orientation-middle", setFormatCellOrientationBorderTopBottomSingleDouble(VerticalAlignment.CENTRE));
        this.format.put("border-tb-single-double-orientation-top", setFormatCellOrientationBorderTopBottomSingleDouble(VerticalAlignment.TOP));
        this.format.put("border-b-double-orientation", setFormatCellOrientationBorderBottomDouble());
        this.format.put("border-b-double-orientation-middle", setFormatCellOrientationBorderBottomDouble(VerticalAlignment.CENTRE));
        this.format.put("border-b-double-orientation-top", setFormatCellOrientationBorderBottomDouble(VerticalAlignment.TOP));
        this.format.put("border-all-b-double-orientation", setFormatCellOrientationBorderAllBottomDouble());
        this.format.put("border-all-b-double-orientation-middle", setFormatCellOrientationBorderAllBottomDouble(VerticalAlignment.CENTRE));
        this.format.put("border-all-b-double-orientation-top", setFormatCellOrientationBorderAllBottomDouble(VerticalAlignment.TOP));

        this.format.put("number-left", setFormatCellNumber(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-center", setFormatCellNumber(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-right", setFormatCellNumber(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-left-top", setFormatCellNumber(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-center-top", setFormatCellNumber(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-right-top", setFormatCellNumber(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-left-middle", setFormatCellNumber(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-center-middle", setFormatCellNumber(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-right-middle", setFormatCellNumber(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));

        this.format.put("number-float-left", setFormatCellNumber(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-center", setFormatCellNumber(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-right", setFormatCellNumber(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-left-top", setFormatCellNumber(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-center-top", setFormatCellNumber(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-right-top", setFormatCellNumber(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-left-middle", setFormatCellNumber(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-center-middle", setFormatCellNumber(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-right-middle", setFormatCellNumber(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));

        this.format.put("number-float-one-left", setFormatCellNumber(Alignment.LEFT, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-center", setFormatCellNumber(Alignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-right", setFormatCellNumber(Alignment.RIGHT, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-left-top", setFormatCellNumber(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-center-top", setFormatCellNumber(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-right-top", setFormatCellNumber(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-left-middle", setFormatCellNumber(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-center-middle", setFormatCellNumber(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-right-middle", setFormatCellNumber(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));

        this.format.put("border-left", setFormatCellBorder(Alignment.LEFT, BorderLineStyle.THIN));
        this.format.put("border-center", setFormatCellBorder(Alignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-right", setFormatCellBorder(Alignment.RIGHT, BorderLineStyle.THIN));
        this.format.put("border-left-top", setFormatCellBorder(Alignment.LEFT, VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-center-top", setFormatCellBorder(Alignment.CENTRE, VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-right-top", setFormatCellBorder(Alignment.RIGHT, VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-left-middle", setFormatCellBorder(Alignment.LEFT, VerticalAlignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-center-middle", setFormatCellBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-right-middle", setFormatCellBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, BorderLineStyle.THIN));

        this.format.put("border-dashed-left", setFormatCellBorder(Alignment.LEFT, BorderLineStyle.DASHED));
        this.format.put("border-dashed-center", setFormatCellBorder(Alignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-dashed-right", setFormatCellBorder(Alignment.RIGHT, BorderLineStyle.DASHED));
        this.format.put("border-dashed-left-top", setFormatCellBorder(Alignment.LEFT, VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-dashed-center-top", setFormatCellBorder(Alignment.CENTRE, VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-dashed-right-top", setFormatCellBorder(Alignment.RIGHT, VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-dashed-left-middle", setFormatCellBorder(Alignment.LEFT, VerticalAlignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-dashed-center-middle", setFormatCellBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-dashed-right-middle", setFormatCellBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, BorderLineStyle.DASHED));

        this.format.put("border-tb-left", setFormatCellBorderTopBottom(Alignment.LEFT, BorderLineStyle.THIN));
        this.format.put("border-tb-center", setFormatCellBorderTopBottom(Alignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-tb-right", setFormatCellBorderTopBottom(Alignment.RIGHT, BorderLineStyle.THIN));
        this.format.put("border-tb-left-top", setFormatCellBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-tb-center-top", setFormatCellBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-tb-right-top", setFormatCellBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, BorderLineStyle.THIN));
        this.format.put("border-tb-left-middle", setFormatCellBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-tb-center-middle", setFormatCellBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, BorderLineStyle.THIN));
        this.format.put("border-tb-right-middle", setFormatCellBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, BorderLineStyle.THIN));

        this.format.put("border-dashed-tb-left", setFormatCellBorderTopBottom(Alignment.LEFT, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-center", setFormatCellBorderTopBottom(Alignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-right", setFormatCellBorderTopBottom(Alignment.RIGHT, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-left-top", setFormatCellBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-center-top", setFormatCellBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-right-top", setFormatCellBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-left-middle", setFormatCellBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-center-middle", setFormatCellBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, BorderLineStyle.DASHED));
        this.format.put("border-dashed-tb-right-middle", setFormatCellBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, BorderLineStyle.DASHED));

        this.format.put("border-tb-double-left", setFormatCellBorderTopBottomDouble(Alignment.LEFT));
        this.format.put("border-tb-double-center", setFormatCellBorderTopBottomDouble(Alignment.CENTRE));
        this.format.put("border-tb-double-right", setFormatCellBorderTopBottomDouble(Alignment.RIGHT));
        this.format.put("border-tb-double-left-top", setFormatCellBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.TOP));
        this.format.put("border-tb-double-center-top", setFormatCellBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP));
        this.format.put("border-tb-double-right-top", setFormatCellBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP));
        this.format.put("border-tb-double-left-middle", setFormatCellBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE));
        this.format.put("border-tb-double-center-middle", setFormatCellBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE));
        this.format.put("border-tb-double-right-middle", setFormatCellBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE));

        this.format.put("border-tb-single-double-left", setFormatCellBorderTopBottomSingleDouble(Alignment.LEFT));
        this.format.put("border-tb-single-double-center", setFormatCellBorderTopBottomSingleDouble(Alignment.CENTRE));
        this.format.put("border-tb-single-double-right", setFormatCellBorderTopBottomSingleDouble(Alignment.RIGHT));
        this.format.put("border-tb-single-double-left-top", setFormatCellBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.TOP));
        this.format.put("border-tb-single-double-center-top", setFormatCellBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.TOP));
        this.format.put("border-tb-single-double-right-top", setFormatCellBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.TOP));
        this.format.put("border-tb-single-double-left-middle", setFormatCellBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.CENTRE));
        this.format.put("border-tb-single-double-center-middle", setFormatCellBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.CENTRE));
        this.format.put("border-tb-single-double-right-middle", setFormatCellBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.CENTRE));

        this.format.put("border-b-double-left", setFormatCellBorderBottomDouble(Alignment.LEFT));
        this.format.put("border-b-double-center", setFormatCellBorderBottomDouble(Alignment.CENTRE));
        this.format.put("border-b-double-right", setFormatCellBorderBottomDouble(Alignment.RIGHT));
        this.format.put("border-b-double-left-top", setFormatCellBorderBottomDouble(Alignment.LEFT, VerticalAlignment.TOP));
        this.format.put("border-b-double-center-top", setFormatCellBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP));
        this.format.put("border-b-double-right-top", setFormatCellBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP));
        this.format.put("border-b-double-left-middle", setFormatCellBorderBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE));
        this.format.put("border-b-double-center-middle", setFormatCellBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE));
        this.format.put("border-b-double-right-middle", setFormatCellBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE));

        this.format.put("border-all-b-double-left", setFormatCellBorderAllBottomDouble(Alignment.LEFT));
        this.format.put("border-all-b-double-center", setFormatCellBorderAllBottomDouble(Alignment.CENTRE));
        this.format.put("border-all-b-double-right", setFormatCellBorderAllBottomDouble(Alignment.RIGHT));
        this.format.put("border-all-b-double-left-top", setFormatCellBorderAllBottomDouble(Alignment.LEFT, VerticalAlignment.TOP));
        this.format.put("border-all-b-double-center-top", setFormatCellBorderAllBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP));
        this.format.put("border-all-b-double-right-top", setFormatCellBorderAllBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP));
        this.format.put("border-all-b-double-left-middle", setFormatCellBorderAllBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE));
        this.format.put("border-all-b-double-center-middle", setFormatCellBorderAllBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE));
        this.format.put("border-all-b-double-right-middle", setFormatCellBorderAllBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE));

        this.format.put("number-border-left", setFormatCellNumberBorder(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-center", setFormatCellNumberBorder(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-right", setFormatCellNumberBorder(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-left-top", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-center-top", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-right-top", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-left-middle", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-center-middle", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-right-middle", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));

        this.format.put("number-border-dashed-left", setFormatCellNumberBorder(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-center", setFormatCellNumberBorder(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-right", setFormatCellNumberBorder(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-left-top", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-center-top", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-right-top", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-left-middle", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-center-middle", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-right-middle", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));

        this.format.put("number-border-tb-left", setFormatCellNumberBorderTopBottom(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-center", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-right", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-left-top", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-center-top", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-right-top", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-left-middle", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-center-middle", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));
        this.format.put("number-border-tb-right-middle", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.THIN));

        this.format.put("number-border-dashed-tb-left", setFormatCellNumberBorderTopBottom(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-center", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-right", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-left-top", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-center-top", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-right-top", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-left-middle", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-center-middle", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));
        this.format.put("number-border-dashed-tb-right-middle", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER, BorderLineStyle.DASHED));

        this.format.put("number-border-tb-double-left", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-center", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-right", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-left-top", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-center-top", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-right-top", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-left-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-center-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-double-right-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));

        this.format.put("number-border-tb-single-double-left", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-center", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-right", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-left-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-center-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-right-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-left-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-center-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-tb-single-double-right-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));

        this.format.put("number-border-b-double-left", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-center", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-right", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-left-top", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-center-top", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-right-top", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-left-middle", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-center-middle", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-b-double-right-middle", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));

        this.format.put("number-border-all-b-double-left", setFormatCellNumberBorderAllBottomDouble(Alignment.LEFT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-center", setFormatCellNumberBorderAllBottomDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-right", setFormatCellNumberBorderAllBottomDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-left-top", setFormatCellNumberBorderAllBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-center-top", setFormatCellNumberBorderAllBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-right-top", setFormatCellNumberBorderAllBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-left-middle", setFormatCellNumberBorderAllBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-center-middle", setFormatCellNumberBorderAllBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));
        this.format.put("number-border-all-b-double-right-middle", setFormatCellNumberBorderAllBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_INTEGER));

        this.format.put("number-float-border-left", setFormatCellNumberBorder(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-center", setFormatCellNumberBorder(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-right", setFormatCellNumberBorder(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-left-top", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-center-top", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-right-top", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-left-middle", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-center-middle", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-right-middle", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));

        this.format.put("number-float-border-dashed-left", setFormatCellNumberBorder(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-center", setFormatCellNumberBorder(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-right", setFormatCellNumberBorder(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-left-top", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-center-top", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-right-top", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-left-middle", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-center-middle", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-right-middle", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));

        this.format.put("number-float-one-border-left", setFormatCellNumberBorder(Alignment.LEFT, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-center", setFormatCellNumberBorder(Alignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-right", setFormatCellNumberBorder(Alignment.RIGHT, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-left-top", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-center-top", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-right-top", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-left-middle", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-center-middle", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-right-middle", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));

        this.format.put("number-float-one-border-dashed-left", setFormatCellNumberBorder(Alignment.LEFT, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-center", setFormatCellNumberBorder(Alignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-right", setFormatCellNumberBorder(Alignment.RIGHT, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-left-top", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-center-top", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-right-top", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-left-middle", setFormatCellNumberBorder(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-center-middle", setFormatCellNumberBorder(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-right-middle", setFormatCellNumberBorder(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));

        this.format.put("number-float-border-tb-left", setFormatCellNumberBorderTopBottom(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-center", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-right", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-left-top", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-center-top", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-right-top", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-left-middle", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-center-middle", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));
        this.format.put("number-float-border-tb-right-middle", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.THIN));

        this.format.put("number-float-border-dashed-tb-left", setFormatCellNumberBorderTopBottom(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-center", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-right", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-left-top", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-center-top", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-right-top", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-left-middle", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-center-middle", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));
        this.format.put("number-float-border-dashed-tb-right-middle", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT, BorderLineStyle.DASHED));

        this.format.put("number-float-one-border-tb-left", setFormatCellNumberBorderTopBottom(Alignment.LEFT, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-center", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-right", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-left-top", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-center-top", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-right-top", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-left-middle", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-center-middle", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));
        this.format.put("number-float-one-border-tb-right-middle", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.THIN));

        this.format.put("number-float-one-border-dashed-tb-left", setFormatCellNumberBorderTopBottom(Alignment.LEFT, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-center", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-right", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-left-top", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-center-top", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-right-top", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-left-middle", setFormatCellNumberBorderTopBottom(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-center-middle", setFormatCellNumberBorderTopBottom(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));
        this.format.put("number-float-one-border-dashed-tb-right-middle", setFormatCellNumberBorderTopBottom(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0"), BorderLineStyle.DASHED));

        this.format.put("number-float-border-tb-double-left", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-center", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-right", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-left-top", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-center-top", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-right-top", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-left-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-center-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-double-right-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));

        this.format.put("number-float-one-border-tb-double-left", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-center", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-right", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-left-top", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-center-top", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-right-top", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-left-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-center-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-double-right-middle", setFormatCellNumberBorderTopBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));

        this.format.put("number-float-border-tb-single-double-left", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-center", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-right", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-left-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-center-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-right-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-left-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-center-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-tb-single-double-right-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));

        this.format.put("number-float-one-border-tb-single-double-center", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-right", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-left-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-center-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-right-top", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-left-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-center-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-tb-single-double-right-middle", setFormatCellNumberBorderTopBottomSingleDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));

        this.format.put("number-float-border-b-double-left", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-center", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-right", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-left-top", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-center-top", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-right-top", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-left-middle", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-center-middle", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-b-double-right-middle", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));

        this.format.put("number-float-border-all-b-double-left", setFormatCellNumberBorderAllBottomDouble(Alignment.LEFT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-center", setFormatCellNumberBorderAllBottomDouble(Alignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-right", setFormatCellNumberBorderAllBottomDouble(Alignment.RIGHT, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-left-top", setFormatCellNumberBorderAllBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-center-top", setFormatCellNumberBorderAllBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-right-top", setFormatCellNumberBorderAllBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-left-middle", setFormatCellNumberBorderAllBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-center-middle", setFormatCellNumberBorderAllBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));
        this.format.put("number-float-border-all-b-double-right-middle", setFormatCellNumberBorderAllBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, NumberFormats.THOUSANDS_FLOAT));

        this.format.put("number-float-one-border-b-double-center", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-right", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-left-top", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-center-top", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-right-top", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.TOP, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-left-middle", setFormatCellNumberBorderBottomDouble(Alignment.LEFT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-center-middle", setFormatCellNumberBorderBottomDouble(Alignment.CENTRE, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
        this.format.put("number-float-one-border-b-double-right-middle", setFormatCellNumberBorderBottomDouble(Alignment.RIGHT, VerticalAlignment.CENTRE, new NumberFormat("#,##0.0")));
    }

    private void setFont() {
        this.font = new WritableFont(this.fontName.get(this.defaultFontName), this.defaultFontSize, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, this.color.get(this.defaultFontColor));
    }

    public double getDefaultFontSize() {
        return defaultFontSize;
    }

    void setDefaultFontSize(int defaultFontSize) {
        this.defaultFontSize = defaultFontSize;
    }

    public String getDefaultFontColor() {
        return this.defaultFontColor;
    }

    void setDefaultFontColor(String defaultFontColor) {
        this.defaultFontColor = defaultFontColor;
    }

    public String getDefaultFontName() {
        return defaultFontName;
    }

    void setDefaultFontName(String defaultFontName) {
        this.defaultFontName = defaultFontName;
    }

    public LinkedHashMap<String, WritableCellFormat> getFormat() {
        return format;
    }

    WritableCellFormat getFormat(String key) {
        if (this.format.containsKey(key)) {
            return this.format.get(key);
        } else {
            printError(key, this.format);
            return this.format.get("border-center");
        }
    }

    private WritableCellFormat setFormatCell(Alignment align) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorder(Alignment align, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setBorder(Border.ALL, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderTopBottom(Alignment align, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setBorder(Border.TOP, borderLineStyle);
            wcf.setBorder(Border.BOTTOM, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderTopBottomDouble(Alignment align) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderTopBottomSingleDouble(Alignment align) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderBottomDouble(Alignment align) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderAllBottomDouble(Alignment align) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientation() {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientation(VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorder(BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setBorder(Border.ALL, borderLineStyle);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorder(VerticalAlignment valign, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.ALL, borderLineStyle);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderTopBottom(BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setBorder(Border.TOP, borderLineStyle);
            wcf.setBorder(Border.BOTTOM, borderLineStyle);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderTopBottom(VerticalAlignment valign, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, borderLineStyle);
            wcf.setBorder(Border.BOTTOM, borderLineStyle);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderTopBottomDouble() {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderTopBottomDouble(VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderTopBottomSingleDouble() {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderTopBottomSingleDouble(VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderBottomDouble() {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderBottomDouble(VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderAllBottomDouble() {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellOrientationBorderAllBottomDouble(VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(Alignment.CENTRE);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setOrientation(Orientation.PLUS_90);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCell(Alignment align, VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorder(Alignment align, VerticalAlignment valign, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.ALL, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderTopBottom(Alignment align, VerticalAlignment valign, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, borderLineStyle);
            wcf.setBorder(Border.BOTTOM, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderTopBottomDouble(Alignment align, VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderTopBottomSingleDouble(Alignment align, VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderBottomDouble(Alignment align, VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellBorderAllBottomDouble(Alignment align, VerticalAlignment valign) {
        try {
            WritableCellFormat wcf = new WritableCellFormat();
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumber(Alignment align, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorder(Alignment align, DisplayFormat format, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setBorder(Border.ALL, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderTopBottom(Alignment align, DisplayFormat format, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setBorder(Border.TOP, borderLineStyle);
            wcf.setBorder(Border.BOTTOM, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderTopBottomDouble(Alignment align, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderTopBottomSingleDouble(Alignment align, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderBottomDouble(Alignment align, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderAllBottomDouble(Alignment align, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumber(Alignment align, VerticalAlignment valign, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorder(Alignment align, VerticalAlignment valign, DisplayFormat format, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.ALL, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderTopBottom(Alignment align, VerticalAlignment valign, DisplayFormat format, BorderLineStyle borderLineStyle) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, borderLineStyle);
            wcf.setBorder(Border.BOTTOM, borderLineStyle);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderTopBottomDouble(Alignment align, VerticalAlignment valign, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, BorderLineStyle.DOUBLE);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderTopBottomSingleDouble(Alignment align, VerticalAlignment valign, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderBottomDouble(Alignment align, VerticalAlignment valign, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private WritableCellFormat setFormatCellNumberBorderAllBottomDouble(Alignment align, VerticalAlignment valign, DisplayFormat format) {
        try {
            WritableCellFormat wcf = new WritableCellFormat(format);
            wcf.setAlignment(align);
            wcf.setVerticalAlignment(valign);
            wcf.setBorder(Border.ALL, BorderLineStyle.THIN);
            wcf.setBorder(Border.BOTTOM, BorderLineStyle.DOUBLE);
            wcf.setFont(this.font);
            return wcf;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private void printError(String msg, Map map) {
        System.out.println("FormatExcelFreedom error => connot found key => " + msg + " in map => " + map);
    }

}