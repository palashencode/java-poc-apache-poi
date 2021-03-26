package com.java.app.excelprocessor;

import javax.net.ssl.SSLEngineResult.HandshakeStatus;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFBorderFormatting;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

public class XSSFFactoryUtil {

    public static final short COLOR_BLACK = IndexedColors.BLACK.getIndex();
    public static final short COLOR_BLUE = IndexedColors.BLUE.getIndex();
    public static final short COLOR_BLUE_GREY = IndexedColors.BLUE_GREY.getIndex();
    public static final short COLOR_BROWN = IndexedColors.BROWN.getIndex();
    public static final short COLOR_GREY_25 = IndexedColors.GREY_25_PERCENT.getIndex();
    public static final String FONT_ARIAL = "Arial";
    public static final String FONT_TAHOMA = "Tahoma";
    public static final Boolean WEIGHT_BOLD = true;
    public static final Boolean WEIGHT_NOBOLD = false;


    public static XSSFSheet createSheet(XSSFWorkbook xssfWorkbook, String name, int order){
        if( xssfWorkbook == null ) return null;
        XSSFSheet sheet = xssfWorkbook.createSheet(name);
        xssfWorkbook.setSheetOrder(name, order);
        return sheet;
    }

    public static XSSFFont createFont(XSSFWorkbook xssfWorkbook, String fontName, int heightInPoints, boolean isBold, short color){
        XSSFFont font = xssfWorkbook.createFont();
        font.setColor(color);
        font.setBold(isBold);
        font.setFontName(fontName);
        font.setFontHeightInPoints((short)heightInPoints);
        return font;
    }

    
    public static XSSFCell addCellDouble(XSSFSheet xssfSheet, int colNo, int rowNo, Double value, XSSFCellStyle xssfCellStyle){
        XSSFCell cell = null;
        cell = xssfSheet.createRow(rowNo).createCell(colNo);
        // cell.setCellType(CellType.BLANK);
        cell.setCellValue(value);
        cell.setCellStyle(xssfCellStyle);
        return cell;
    }

    public static XSSFCell addCellString(XSSFSheet xssfSheet, int colNo, int rowNo, String value, XSSFCellStyle xssfCellStyle){
        XSSFCell cell = null;
        cell = xssfSheet.createRow(rowNo).createCell(colNo);
        // cell.setCellType(CellType.BLANK);
        cell.setCellValue(value);
        cell.setCellStyle(xssfCellStyle);
        return cell;
    }

    public static XSSFCell addNewHyperLinkStringCell(XSSFSheet xssfSheet, int colNo, int rowNo, String label, String url, XSSFCellStyle xssfCellStyle){
        
        XSSFCell cell = XSSFFactoryUtil.addCellString(xssfSheet, colNo, rowNo, label, xssfCellStyle);
        
        CreationHelper creationHelper = xssfSheet.getWorkbook().getCreationHelper();
        Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress(url);
        cell.setHyperlink(link);
        return cell;
    }

    public static XSSFCell addLinkToCell(XSSFWorkbook xssfWorkbook, XSSFCell cell, String url){
        CreationHelper creationHelper = xssfWorkbook.getCreationHelper();
        Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress(url);
        cell.setHyperlink(link);
        return cell;
    }

    public static XSSFCell addCellNumber(XSSFSheet xssfSheet, int colNo, int rowNo, Integer value, XSSFCellStyle xssfCellStyle){
        XSSFCell cell = null;
        cell = xssfSheet.createRow(rowNo).createCell(colNo);
        // cell.setCellType(CellType.BLANK);
        cell.setCellValue(value);
        cell.setCellStyle(xssfCellStyle);
        return cell;
    }

    public static XSSFCellStyle createCellStyle(XSSFWorkbook xssfWorkbook){
            XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
            return xssfCellStyle;
    }

    public static XSSFCellStyle createCellStyle(XSSFWorkbook xssfWorkbook, XSSFFont xssfFont){
        XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
        xssfCellStyle.setFont(xssfFont);
        return xssfCellStyle;
    }

    public static XSSFCellStyle createCellStyleCustomNumberFormat(XSSFWorkbook xssfWorkbook, XSSFFont xssfFont, String dataFormat){
        XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
        xssfCellStyle.setFont(xssfFont);
        xssfCellStyle.setDataFormat(xssfWorkbook.createDataFormat().getFormat(dataFormat));
        return xssfCellStyle;
    }

    public static XSSFCellStyle createCellStyle(XSSFWorkbook xssfWorkbook, XSSFFont xssfFont, HorizontalAlignment hAlignment){
        XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
        xssfCellStyle.setFont(xssfFont);
        xssfCellStyle.setAlignment(hAlignment);
        return xssfCellStyle;
    }

    public static XSSFCellStyle createCellStyle(XSSFWorkbook xssfWorkbook, XSSFFont xssfFont, short colorIndex){
        XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
        xssfCellStyle.setFont(xssfFont);
        setBackGroundColor(xssfCellStyle,colorIndex);
        return xssfCellStyle;
    }

        public static XSSFCellStyle createCellStyle(XSSFWorkbook xssfWorkbook, XSSFFont xssfFont, short colorIndex
                                                , HorizontalAlignment hAlignment){
        XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
        
        xssfCellStyle.setFont(xssfFont);
        setBackGroundColor(xssfCellStyle,colorIndex);
        xssfCellStyle.setAlignment(hAlignment);
        return xssfCellStyle;
    }

    private static void setBackGroundColor(XSSFCellStyle style, int colorIndex){
        style.setFillForegroundColor((short)colorIndex);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public static XSSFCellStyle createCellStyle(XSSFWorkbook xssfWorkbook, XSSFFont xssfFont, short colorIndex
                                                , HorizontalAlignment hAlignment, VerticalAlignment vAlignment){
        XSSFCellStyle xssfCellStyle = xssfWorkbook.createCellStyle();
        xssfCellStyle.setFont(xssfFont);
        setBackGroundColor(xssfCellStyle,colorIndex);
        xssfCellStyle.setAlignment(hAlignment);
        xssfCellStyle.setVerticalAlignment(vAlignment);
        return xssfCellStyle;
    }

    public static void setBorderStyle(XSSFCellStyle xssfCellStyle,short colorIndex, BorderStyle borderStyle, boolean left, boolean right, boolean top, boolean bottom){
            if(left){
                xssfCellStyle.setLeftBorderColor(colorIndex);
                xssfCellStyle.setBorderLeft(borderStyle);
            }
            if(right){
                xssfCellStyle.setRightBorderColor(colorIndex);
                xssfCellStyle.setBorderRight(borderStyle);
            }
            if(top){
                xssfCellStyle.setTopBorderColor(colorIndex);
                xssfCellStyle.setBorderTop(borderStyle);
            }
            if(bottom){
                xssfCellStyle.setBottomBorderColor(colorIndex);
                xssfCellStyle.setBorderBottom(borderStyle);
            }
    }

    public static void newBorderStyle(XSSFWorkbook workbook, XSSFCellStyle oldCellStyle,short colorIndex, BorderStyle borderStyle, boolean left, boolean right, boolean top, boolean bottom){
        XSSFCellStyle xssfCellStyle = workbook.createCellStyle();
        xssfCellStyle.cloneStyleFrom(oldCellStyle);

        if(left){
            xssfCellStyle.setLeftBorderColor(colorIndex);
            xssfCellStyle.setBorderLeft(borderStyle);
        }
        if(right){
            xssfCellStyle.setRightBorderColor(colorIndex);
            xssfCellStyle.setBorderRight(borderStyle);
        }
        if(top){
            xssfCellStyle.setTopBorderColor(colorIndex);
            xssfCellStyle.setBorderTop(borderStyle);
        }
        if(bottom){
            xssfCellStyle.setBottomBorderColor(colorIndex);
            xssfCellStyle.setBorderBottom(borderStyle);
        }
}

}
