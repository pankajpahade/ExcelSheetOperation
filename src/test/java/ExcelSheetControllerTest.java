import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.jupiter.api.Test;
import org.junit.runner.RunWith;
import org.junit.validator.AnnotationsValidator;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.MockitoAnnotations;
import org.mockito.runners.MockitoJUnitRunner;

import java.io.IOException;

import static junit.framework.Assert.assertEquals;
import static org.mockito.Mockito.when;

//@RunWith(MockitoJUnitRunner.class)
class ExcelSheetControllerTest {

    @Mock
    public  ExcelSheetService service = Mockito.mock(ExcelSheetService.class);

    @InjectMocks
    public ExcelSheetController controller = new ExcelSheetController(service);

    @Before
    public void setUp() {
        MockitoAnnotations.initMocks(this);
        //service = new ExcelSheetService();
        //service.writeExcelSheet();
    }

    @Test
    public void writeExcelSheet() {
        when(service.writeExcelSheet()).thenReturn("EmpOjas.xls is Successfully Created");
        String act = controller.createExcelSheetWithData();
        String expect = "EmpOjas.xls is Successfully Created";
        assertEquals(expect, act);
    }

 /* @Test
    public void copyExcelSheetIntoNew() throws IOException {

        when(service.copyExcelSheetIntoNew()).thenReturn();

    }*/

    @Test
    public void readAndStoreDataIntoDB() {

    }

    @Test
    public void checkExcelSheetEmptyOrNot() throws IOException {
        when(service.isExcelSheetEmpty()).thenReturn(Boolean.valueOf("Excel Sheet is Empty"));
        String actual = controller.checkExcelSheetEmptyOrNot();
        String expected = "Excel Sheet is Empty";
        assertEquals(expected,actual);
    }
}