package Windchill;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

//import org.apache.commons.compress.utils.IOUtils;

public class OutputWrite {
    FileOutputStream fileOut;

    public void WriteFileContent(String[] Data, int row, String file) throws IOException, InvalidFormatException {
        XSSFWorkbook workbook;
        XSSFSheet sheet;

        if (row == 0) {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Test Results");
            XSSFCellStyle headerstyle = workbook.createCellStyle();
            headerstyle = setCellStyle(workbook, "header", headerstyle);
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Product Name");
            header.createCell(1).setCellValue("Part Name");
            header.createCell(2).setCellValue("General Details");
            for (int j = 0; j <= 2; j++) {
                header.getCell(j).setCellStyle(headerstyle);
            }
            Row pRow = sheet.createRow(1);
            Cell cell;
            for (int i = 0; i < Data.length; i++) {
                cell = pRow.createCell(i);
                cell.setCellValue(Data[i]);
            }
            for (int i = 0; i < Data.length + 1; i++) {
                sheet.autoSizeColumn(i);
            }
            XSSFCellStyle style = workbook.createCellStyle();
            style = setCellStyle(workbook, "NotMatched", style);
        } else {
            FileInputStream inputStream = new FileInputStream(new File(file));
            workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);
            sheet = workbook.getSheet("Test Results");
            Row pRow = sheet.createRow(row + 1);
            Cell cell ;
            for (int i = 0; i < Data.length; i++) {
                cell = pRow.createCell(i);
                cell.setCellValue(Data[i]);
            }
            for (int i = 0; i <  Data.length + 1; i++) {
                sheet.autoSizeColumn(i);
            }
            XSSFCellStyle style = workbook.createCellStyle();
            style = setCellStyle(workbook, "NotMatched", style);
        }
        fileOut = new FileOutputStream(file);
        workbook.write(fileOut);
        //workbook.close();
    }

    public void addPicture(Workbook workbook, XSSFSheet sheet, String[] Data, int row) throws IOException {
        InputStream is = new FileInputStream(Data[7]);
        byte[] bytes = IOUtils.toByteArray(is);
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        is.close();
        XSSFCreationHelper helper = (XSSFCreationHelper) workbook.getCreationHelper();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = helper.createClientAnchor();
        anchor = helper.createClientAnchor();
        anchor.setCol1(9);
        anchor.setCol2(10);
        anchor.setRow1(row + 1);
        anchor.setRow2(row + 2);
        drawing.createPicture(anchor, pictureIdx);
    }

    public String createFileName() {
        LocalDateTime localDateTime = LocalDateTime.now();
        DateTimeFormatter formatter1 = DateTimeFormatter.ofPattern("dd-LLLL-yyyy");
        DateTimeFormatter formatter2 = DateTimeFormatter.ofPattern("HH.mm.ss");
        String date = localDateTime.format(formatter1);
        String time = localDateTime.format(formatter2);
        String directory = "Results\\Results_" + date;
        String filename = "\\result_" + time + ".xlsx";
        File dir = new File(directory);
        if (!dir.exists())
            dir.mkdirs();
        return directory + "/" + filename;
    }

    public XSSFCellStyle setCellStyle(Workbook workbook, String CellType, XSSFCellStyle style) {
        if (CellType == "header") {
            Font font = workbook.createFont();
            font.setFontName(XSSFFont.DEFAULT_FONT_NAME);
            font.setFontHeightInPoints((short) 10);
            font.setBoldweight((short) 10);
            //font.setBold(true);
            font.setColor(IndexedColors.WHITE.index);
            style.setFont(font);
            style.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
            style.setFillPattern(FillPatternType.FINE_DOTS);
        } /*
         * else if (CellType == "NotMatched") {
         * style.setFillBackgroundColor(IndexedColors.RED.index);
         * style.setBottomBorderColor(IndexedColors.BLACK.index);
         * style.setFillPattern(FillPatternType.FINE_DOTS); }
         */
        return style;
    }

    public void CloseandSave() throws IOException {
        System.out.println("EndofTesting");
        fileOut.close();
    }
}