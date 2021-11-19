package Windchill;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ExcelReader {

    private FileInputStream inputStream;
    private FileOutputStream fos;
    private Workbook workbook;
    private Sheet worksheet;
    private HashMap<String, List<String>> tableData = new HashMap<String, List<String>>();

    private ExcelReader(String filePath, String fileName, String sheetName) {
        File file = new File(filePath + "\\" + fileName);
        try {
            inputStream = new FileInputStream(file);
            String fileExtensionName = fileName.substring(fileName.lastIndexOf("."));

            if (fileExtensionName.equals(".xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            }else if(fileExtensionName.equals(".xls")){
                workbook = new HSSFWorkbook(inputStream);
            }
            worksheet = workbook.getSheet(sheetName);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public Sheet getWorksheet(String globalSettings) {
        return worksheet;
    }

    public void setWorksheet(Sheet worksheet) {
        this.worksheet = worksheet;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /*public static ExcelReader getInstance(String filePath, String fileName, String sheetName) {
        return new ExcelReader(filePath, fileName, sheetName);
    }*/

    public static Windchill.ExcelReader getInstance(String filePath, String fileName, String sheetName) {
        Windchill.ExcelReader excelReader = new Windchill.ExcelReader(filePath, fileName, sheetName);
        return excelReader;
    }


    public void WriteData(String CompletePath, String sheetName,String Columnname,String productno) throws IOException {
        FileInputStream inputStream;
        FileOutputStream fos;
        fos=null;
        inputStream=new FileInputStream(CompletePath);
        XSSFWorkbook workbook1=new XSSFWorkbook(inputStream);
        XSSFSheet sheet=workbook1.getSheet(sheetName);
        XSSFRow row=null;
        XSSFCell cell=null;
        int colnum=-1;
        row=sheet.getRow(0);
        for(int i=0;i<row.getLastCellNum();i++)
        {
            // System.out.println("1:"+row.getCell(i).getStringCellValue());
            //System.out.println("2:"+Columnname );
            if(row.getCell(i).getStringCellValue().equals(Columnname))
            {
                //System.out.println(row.getCell(i).getStringCellValue());
                colnum=i;
            }
        }
        row=sheet.getRow(1);
        if(row==null)
            row=sheet.createRow(1);
        cell=row.getCell(colnum);
        if(cell==null)
            cell=row.createCell(colnum);
        cell.setCellValue(productno);
        fos=new FileOutputStream(CompletePath);
        workbook1.write(fos);
        fos.close();
    }

    public String WriteData(String CompletePath, String sheetName, Integer rownumber, String Columnname, String productno) throws IOException {
        FileInputStream inputStream;
        FileOutputStream fos;
        fos=null;
        inputStream=new FileInputStream(CompletePath);
        XSSFWorkbook workbook2=new XSSFWorkbook(inputStream);
        XSSFSheet sheet=workbook2.getSheet(sheetName);
        XSSFRow row=null;
        XSSFCell cell=null;
        int colnum=-1;
        row=sheet.getRow(0);
        for(int i=0;i<row.getLastCellNum();i++)
        {
            // System.out.println("1:"+row.getCell(i).getStringCellValue());
            //System.out.println("2:"+Columnname );
            if(row.getCell(i).getStringCellValue().equals(Columnname))
            {
                //  System.out.println(row.getCell(i).getStringCellValue());
                colnum=i;
            }
        }
        row=sheet.getRow(rownumber);
        if(row==null)
            row=sheet.createRow(rownumber);
        cell=row.getCell(colnum);
        if(cell==null)
            cell=row.createCell(colnum);
        cell.setCellValue(productno);
        fos=new FileOutputStream(CompletePath);
        workbook2.write(fos);
        fos.close();
        return CompletePath;
    }
    public String RetrieveData(String CompletePath, String Sheetname, int row, int column) throws IOException {
        inputStream=new FileInputStream(CompletePath);
        XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
        worksheet=workbook.getSheet(Sheetname);
        String data= worksheet.getRow(row).getCell(column).getStringCellValue();
        return data;
    }

    public String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString();
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double doubleValue = cell.getNumericCellValue();
                    double intValue = (double)((int) doubleValue);
                    if(intValue==doubleValue) {
                        return Integer.toString((int)cell.getNumericCellValue());
                    }else {
                        return Double.toString(doubleValue);
                    }
                }
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case Cell.CELL_TYPE_BLANK:
                return "";
            default:
                return "";
        }
    }

    public List<String> getRowData(int rowIndex, int columnStartIndex) {
        Row row = worksheet.getRow(rowIndex);
        List<String> rowData = new ArrayList<String>();
        if (columnStartIndex < row.getLastCellNum()) {
            for (int j = columnStartIndex; j < row.getLastCellNum(); j++) {
                String value = getCellValue(row.getCell(j));
                if(value!=null && value!="") {
                    rowData.add(value);
                }
            }
        }
        return rowData;
    }
}