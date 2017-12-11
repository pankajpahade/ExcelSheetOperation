import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.*;
import java.util.*;

public class ExcelSheetController {

    @Autowired
    private ExcelSheetService service;

    public ExcelSheetController() {
    }

    public ExcelSheetController(ExcelSheetService service) {
        this.service = service;
    }

    public String createExcelSheetWithData(){
        return service.writeExcelSheet();
    }

    public void insertRecordInAlreadyCreatedExcelSheet(){
        service.insertNewRecordExcelSheet();
    }

    public void createNewExcelSheetAndAddDataFromAlreadyCreatedSheet(){
        service.copyExcelSheetIntoNew();
    }

    public void readAndStoreDataIntoDB(){
        ExcelSheetService service = new ExcelSheetService();
        String fileName = "C:/Sample/ExcelFileOperation/EmpOjas.xlsx";
        Vector dataToSave = service.readDataFromExcelSheet(fileName);
        service.saveExcelSheetDataToDB(dataToSave);
    }

    public String checkExcelSheetEmptyOrNot(){
        ExcelSheetService service = new ExcelSheetService();
        String result = null;
        try {
            if(service.isExcelSheetEmpty())
                result = "Excel Sheet is not Empty";
            else
                result = "Excel Sheet is Empty";
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    public void getDuplicatesFromExcellSheet(){
        try {
            service.getDuplicates();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void getSizeOfExcellSheet(){
        try {
            service.sizeOfTheExcelSheet();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}






