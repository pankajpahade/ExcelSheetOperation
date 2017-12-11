import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.*;

public class ExcelSheetService {

    public String writeExcelSheet() {

        try {
            FileOutputStream outputStream = new FileOutputStream(new File("EmpOjas.xlsx"));

            //Blank workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            //Create a blank sheet
            XSSFSheet sheet = workbook.createSheet("EmpOjas.xlsx");

            Map employeeData = new HashMap();

            employeeData.put("1", new Object[]{"ID", "FirstName", "LastName"});
            employeeData.put("2", new Object[]{"101", "Pankaj", "Pahade"});
            employeeData.put("3", new Object[]{"102", "Ayyappa", "Sigatapu"});
            employeeData.put("4", new Object[]{"103", "Farid", "Basha"});

            Set keySet = employeeData.keySet();

            int rowNum = 0;
            for (Object key : keySet) {

                Row row = sheet.createRow(rowNum++);
                Object[] objectArray = (Object[]) employeeData.get(key);
                int cellNum = 0;

                for (Object object : objectArray) {

                    Cell cell = row.createCell(cellNum++);

                    if (object instanceof String) {
                        cell.setCellValue((String) object);
                    } else if (object instanceof Integer) {
                        cell.setCellValue((Integer) object);
                    }
                }
            }
            workbook.write(outputStream);
            outputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
        return "EmpOjas.xls is Successfully Created";

    }

    public String copyExcelSheetIntoNew() {

        try {
            InputStream excelFileToRead = new FileInputStream("EmpOjas.xlsx");
            OutputStream newExcelSheetForCopy = new FileOutputStream("EmpOjasCopy.xlsx");

            XSSFWorkbook wb = new XSSFWorkbook(excelFileToRead);
            XSSFWorkbook newWb = new XSSFWorkbook();

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFSheet newSheet = newWb.createSheet("EmpOjasCopy.xlsx");

            Iterator rows = sheet.rowIterator();
            int rowNum = 0;

            while (rows.hasNext()) {
                Row row = (XSSFRow) rows.next();
                Row newRow = newSheet.createRow(rowNum++);

                Iterator cells = row.cellIterator();
                int cellNum=0;

                while (cells.hasNext() ) {

                    Cell cell = (XSSFCell) cells.next();
                    Cell newCell = newRow.createCell(cellNum++);

                    if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        newCell.setCellValue(cell.getStringCellValue());
                    } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                        newCell.setCellValue(cell.getNumericCellValue());
                    }
                }
            }
            newWb.write(newExcelSheetForCopy);

            excelFileToRead.close();
            newExcelSheetForCopy.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "EmpOjas.xls is Copyed into EmpOjasCopy.xls Successfully";
    }

    public boolean  isExcelSheetEmpty() throws IOException {
        InputStream inputStream = new FileInputStream(new File("EmpOjas.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()) {
            Row row = (Row) rows.next();
            Iterator cells = row.cellIterator();
           if(cells.hasNext()) {
               return true;
            }
        }
        return false;
    }

    public Vector readDataFromExcelSheet(String fileName){

        Vector cellVectorHolder = new Vector();

        try {
            FileInputStream inputStream = new FileInputStream(fileName);
            XSSFWorkbook workbook= new XSSFWorkbook(inputStream);
            XSSFSheet xssfSheet = workbook.getSheetAt(0);
            Iterator rowIterator = xssfSheet.rowIterator();

            while (rowIterator.hasNext()){
                Row row = (Row) rowIterator.next();
                Iterator cellIterator = row.cellIterator();

                Vector cellVector = new Vector();

                while (cellIterator.hasNext()){
                    Cell cell = (Cell) cellIterator.next();
                    cellVector.addElement(cell);
                }
                cellVectorHolder.addElement(cellVector);
            }
        }catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return cellVectorHolder;
    }

    public void saveExcelSheetDataToDB(Vector dataHolder) {
        String ID = "";
        String FirstName = "";
        String LastName = "";

        for (Iterator iterator = dataHolder.iterator(); iterator.hasNext(); ) {
            List list = (List) iterator.next();
            ID = list.get(0).toString();
            FirstName = list.get(1).toString();
            LastName = list.get(2).toString();

            try {
                Class.forName("com.mysql.jdbc.Driver").newInstance();
                Connection con = DriverManager.getConnection("jdbc:mysql://192.168.1.159/pankaj", "root", "root");

                System.out.println("connection implement Successfully...");

                PreparedStatement stmt = con.prepareStatement("INSERT INTO empojas(ID,FirstName,LastName) VALUES(?,?,?)");
                stmt.setString(1, ID);
                stmt.setString(2, FirstName);
                stmt.setString(3, LastName);
                stmt.executeUpdate();

                System.out.println("Data is inserted Into the DataBase Successfully...");

                stmt.close();
                con.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public void sizeOfTheExcelSheet() throws IOException {
        File file = new File("EmpOjas.xlsx");
        double inBytes = file.length();
        double inKiloBytes = (inBytes/1024);
        double inMegaBytes = (inKiloBytes/1024);

        System.out.println("Size of "+file+" file in Bytes is "+inBytes+" Bytes");
        System.out.println("Size of "+file+" file in KB is "+inKiloBytes+" KB");
        System.out.println("Size of "+file+" file in MB is "+inMegaBytes+" MB");
    }

    public void getDuplicates() throws DuplicateNotAllowedException, IOException {
        FileInputStream inputStream = new FileInputStream("EmpOjasCopy.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        Set set = new HashSet();
        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()){
            Row row = (Row) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()){
                Cell cell = (Cell) cells.next();
                if (!set.add(cell.getNumericCellValue())){
                    Object cellValue = null;
                    List listDuplicats = new ArrayList();
                    if(!set.add(cell)) {
                        cellValue = cell.getStringCellValue();
                        listDuplicats.add(cellValue);
                    }
                    else if (!set.add(cell)) {
                        cellValue = cell.getNumericCellValue();
                        listDuplicats.add(cellValue);
                    }
                    System.out.println(listDuplicats);
                    //throw new DuplicateNotAllowedException("Element Already Exists ");
                }
            }
        }
    }

    public void insertNewRecordExcelSheet() {
        try {
            InputStream inputStream = new FileInputStream("EmpOjasCopy.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("EmpOjasCopy.xlsx");
            int lastRow = sheet.getLastRowNum() + 2;
            //Row row = sheet.createRow(lastRow);
            Map employeeData = new HashMap();
            employeeData.put(lastRow, new Object[]{"104", "Durgesh", "Kumar"});
            employeeData.put(lastRow+1, new Object[]{"105", "jhgfd", "jgf"});

            Set keySet = employeeData.keySet();
            // int lastRow = sheet.getLastRowNum();
            //Iterator rows  = sheet.rowIterator();

            for (Object key : keySet){
                Row rows = sheet.createRow(lastRow++);
                Object[] objects = (Object[]) employeeData.get(key);
                // XSSFRow cells = sheet.createRow(lastRow++);
                int cellNum = 0;

                for (Object object : objects){
                    Cell cell = rows.createCell(cellNum++);

                    if(object instanceof String){
                        System.out.println(object.toString());
                        cell.setCellValue((String) object);
                    }else if (object instanceof Integer){
                        cell.setCellValue((Integer) object);
                    }
                }
            }
            System.out.println("Data inserted Successfully");
            inputStream.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
