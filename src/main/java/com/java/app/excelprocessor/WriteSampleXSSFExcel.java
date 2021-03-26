package com.java.app.excelprocessor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;


import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteSampleXSSFExcel {
    public static void main(String args[]){
        System.out.println("Hello from WriteSampleXSSFExcel");
        generateSampleXSSFExcel("./excel/XLSXExample.xlsx");
    }

    private static final Map<String,XSSFCellStyle> styleMap = new HashMap<>();
    private static final Map<String,XSSFFont> fontMap = new HashMap<>();

    private static void generateSampleXSSFExcel(String fileName){
        
        try(XSSFWorkbook wb = new XSSFWorkbook()){
            FileOutputStream fos = new FileOutputStream(fileName);
            XSSFSheet sheet = XSSFFactoryUtil.createSheet(wb, "Cell Types", 0);

            XSSFCellStyle primaryHeaderStyle = getPrimaryHeaderStyle(wb);
            styleMap.put("primary_header", primaryHeaderStyle);
            
            XSSFCellStyle secondaryHeaderStyle = getSecondaryStyle(wb);
            styleMap.put("secondary_header", secondaryHeaderStyle);

            XSSFCellStyle secondaryHeaderRightAlignStyle = getSecondaryStyleRightAlign(wb);
            styleMap.put("secondary_header_right_align", secondaryHeaderRightAlignStyle);
            
            XSSFCellStyle defaultTextStyle = getDefaultTextStyle(wb);
            styleMap.put("default_text", defaultTextStyle);

            XSSFCellStyle defaultTextStyleBold = getDefaultTextStyleBold(wb);
            styleMap.put("default_text_bold", defaultTextStyleBold);

            XSSFCellStyle defaultTahomaTextStyle = getDefaultTahomaTextStyle(wb);
            styleMap.put("default_tahoma_text", defaultTahomaTextStyle);
            
            XSSFCellStyle defaultTahomaTextStyleCurrency = getDefaultTahomaTextStyleCurrency(wb);
            styleMap.put("default_tahoma_text_currency", defaultTahomaTextStyleCurrency);

            XSSFCellStyle defaultTahomaTextStyleInteger = getDefaultTahomaTextStyleInteger(wb);
            styleMap.put("default_tahoma_text_integer", defaultTahomaTextStyleInteger);

            XSSFCellStyle defaultTahomaTextStyleDouble = getDefaultTahomaTextStyleDouble(wb);
            styleMap.put("default_tahoma_text_double", defaultTahomaTextStyleDouble);

            XSSFCellStyle defaultTahomaTextStyleBorder_left_right = getDefaultTahomaTextStyleBorderLeftRight(wb);
            styleMap.put("default_tahoma_text_border_left_right", defaultTahomaTextStyleBorder_left_right);

            XSSFCellStyle hyperlinkStyle = getHyperLinkStyle(wb);
            styleMap.put("hyperlink", hyperlinkStyle);

            XSSFCellStyle labelTextStyle = getLabelCellStyle(wb);
            styleMap.put("label_text", labelTextStyle);

            int row = 2;
            XSSFCell cell1 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Primary Header Style", primaryHeaderStyle); row += 2;
            
            XSSFCell cell2 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Primary Header Style Merged", primaryHeaderStyle);
            sheet.addMergedRegion(new CellRangeAddress(row,row,0,2)); row += 2;

            XSSFCell cell3 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Secondary Header Style", secondaryHeaderStyle); row += 2;
            
            XSSFCell cell31 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Right Align", secondaryHeaderRightAlignStyle); row += 2;

            XSSFCell cell4 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Default Text Style", defaultTextStyle); row += 2;

            XSSFCell cell41 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Default Text Style Bold", defaultTextStyleBold); row += 2;

            XSSFCell cell42 = XSSFFactoryUtil.addCellString(sheet, 2, row, "Default Text Style Border", defaultTahomaTextStyleBorder_left_right); row += 2;

            XSSFCell cell5 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Label Text Style", labelTextStyle); row += 2;

            XSSFCell cell6 = XSSFFactoryUtil.addCellString(sheet, 1, row, 23+"", defaultTahomaTextStyle); row += 2;

            XSSFCell cell7 = XSSFFactoryUtil.addCellNumber(sheet, 1, row, 23, defaultTahomaTextStyle); row += 2;

            XSSFCell cell71 = XSSFFactoryUtil.addCellNumber(sheet, 1, row, 84, defaultTahomaTextStyleCurrency); row += 2;

            XSSFCell cell72 = XSSFFactoryUtil.addCellDouble(sheet, 1, row, 1284.25416, defaultTahomaTextStyleDouble); row += 2;

            XSSFCell cell73 = XSSFFactoryUtil.addCellNumber(sheet, 1, row, 1284, defaultTahomaTextStyleInteger); row += 2;

            XSSFCell cell74 = XSSFFactoryUtil.addCellString(sheet, 0, row, "Secret Link", hyperlinkStyle); row += 2;
            XSSFFactoryUtil.addLinkToCell(sheet.getWorkbook(), cell74, "https://www.google.com/search?q=apache+poi");
        
            XSSFCell cell77 = XSSFFactoryUtil.addNewHyperLinkStringCell(sheet, 0, row, "Secret Link2","https://www.google.com/search?q=apache+poi+again", hyperlinkStyle); row += 2;
            

            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(2);

            wb.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static XSSFCellStyle getLabelCellStyle(XSSFWorkbook xssfWorkbook){
    XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 12, XSSFFactoryUtil.WEIGHT_BOLD, XSSFFactoryUtil.COLOR_BROWN);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        return style;
    }

    private static XSSFCellStyle getPrimaryHeaderStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_TAHOMA, 10, XSSFFactoryUtil.WEIGHT_BOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font, XSSFFactoryUtil.COLOR_BLUE_GREY, HorizontalAlignment.CENTER);
        return style;
    }

    private static XSSFCellStyle getSecondaryStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_TAHOMA, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        return style;
    }

    private static XSSFCellStyle getSecondaryStyleRightAlign(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_TAHOMA, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font,HorizontalAlignment.RIGHT);
        return style;
    }

    private static XSSFCellStyle getDefaultTahomaTextStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        return style;
    }

    private static XSSFCellStyle getDefaultTahomaTextStyleCurrency(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyleCustomNumberFormat(xssfWorkbook, font, "[$$-409]#,##0.00");
        return style;
    }

    private static XSSFCellStyle getDefaultTahomaTextStyleDouble(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyleCustomNumberFormat(xssfWorkbook, font, "#,##0.00");
        return style;
    }

    private static XSSFCellStyle getDefaultTahomaTextStyleInteger(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyleCustomNumberFormat(xssfWorkbook, font, "#,##");
        return style;
    }

    private static XSSFCellStyle getDefaultTahomaTextStyleBorderLeftRight(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 10, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        XSSFFactoryUtil.setBorderStyle(style, XSSFFactoryUtil.COLOR_BLACK, BorderStyle.THIN, true, true, false, false);
        return style;
    }

    private static XSSFCellStyle getDefaultTextStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 8, XSSFFactoryUtil.WEIGHT_NOBOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        return style;
    }

    private static XSSFCellStyle getDefaultTextStyleBold(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_ARIAL, 8, XSSFFactoryUtil.WEIGHT_BOLD, XSSFFactoryUtil.COLOR_BLACK);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        return style;
    }

    private static XSSFCellStyle getHyperLinkStyle(XSSFWorkbook xssfWorkbook){
        XSSFCellStyle style = null;
        XSSFFont font = XSSFFactoryUtil.createFont(xssfWorkbook, XSSFFactoryUtil.FONT_TAHOMA, 10, XSSFFactoryUtil.WEIGHT_BOLD, XSSFFactoryUtil.COLOR_BLUE);
        font.setUnderline(FontUnderline.SINGLE);
        style = XSSFFactoryUtil.createCellStyle(xssfWorkbook, font);
        return style;
    }

    

}
