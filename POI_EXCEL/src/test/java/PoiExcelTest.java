
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.CreationHelper;
//import org.apache.poi.ss.usermodel.DataFormatter;
//import org.apache.poi.ss.usermodel.DateUtil;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author USER
 */
public class PoiExcelTest {

    Workbook wb;
    Sheet sheet;
    Sheet sheet2;
    CreationHelper createHelper;
    String file_name = "C:/test/workbook.xlsx";

    public PoiExcelTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
        wb = new XSSFWorkbook();
        createHelper = wb.getCreationHelper();
        sheet = wb.createSheet("new sheet");
        sheet2 = wb.createSheet("second sheet");
    }

    @After
    public void tearDown() {
    }

    // TODO add test methods here.
    // The methods must be annotated with annotation @Test. For example:
    //
    @Test
    public void testSheet() {

    }

    @Test
    public void testCell() {

        // Create a cell and put a value in it.
        Cell cell = createRow().createCell(0);
        cell.setCellValue(1);

        // Or do it on one line.
        createRow().createCell(1).setCellValue(1.2);
        createRow().createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        createRow().createCell(3).setCellValue(true);
    }

    private Row createRow() {
        return sheet.createRow(0);
    }

    @Test
    public void createDateCell() {
        Row row = sheet.createRow(0);

        // Create a cell and put a date value in it.  The first cell is not styled
        // as a date.
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());

        // we style the second cell as a date (and time).  It is important to
        // create a new cell style from the workbook otherwise you can end up
        // modifying the built in style and effecting not only this cell but other cells.
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        //you can also set date as java.util.Calendar
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

    }

    @Test
    public void differentKindOfCell() {
        createRow().createCell(0).setCellValue(1.1);
        createRow().createCell(1).setCellValue(new Date());
        createRow().createCell(2).setCellValue(Calendar.getInstance());
        createRow().createCell(3).setCellValue("a string");
        createRow().createCell(4).setCellValue(true);
        createRow().createCell(5).setCellType(CellType.ERROR);
    }

    @Test
    public void getCellContent() {
        DataFormatter formatter = new DataFormatter();
        Sheet sheet1 = wb.getSheetAt(0);
        for (Row row : sheet1) {
            for (Cell cell : row) {
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");

                // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
                String text = formatter.formatCellValue(cell);
                System.out.println(text);

                // Alternatively, get the value and format it yourself
//                switch (cell.getCellType()) {
//                    case CellType.STRING:
//                        System.out.println(cell.getRichStringCellValue().getString());
//                        break;
//                    case CellType.NUMERIC:
//                        if (DateUtil.isCellDateFormatted(cell)) {
//                            System.out.println(cell.getDateCellValue());
//                        } else {
//                            System.out.println(cell.getNumericCellValue());
//                        }
//                        break;
//                    case CellType.BOOLEAN:
//                        System.out.println(cell.getBooleanCellValue());
//                        break;
//                    case CellType.FORMULA:
//                        System.out.println(cell.getCellFormula());
//                        break;
//                    case CellType.BLANK:
//                        System.out.println();
//                        break;
//                    default:
//                        System.out.println();
//                }
            }
        }
    }

    @Test
    public void fillColors() {

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);

        // Aqua background
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        Cell cell = row.createCell(1);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Orange "foreground", foreground being the fill foreground not the font color.
        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(2);
        cell.setCellValue("X");
        cell.setCellStyle(style);
    }

    @Test
    public void mergCell() {

        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of merging");

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row  (0-based)
                1, //first column (0-based)
                2 //last column  (0-based)
        ));

    }

    @Test
    public void testWorkBook() throws FileNotFoundException, IOException {
//        testCell();
//        createDateCell();
//        differentKindOfCell();
//        fillColors();
        mergCell();
        try (OutputStream fileOut = new FileOutputStream(file_name)) {
            wb.write(fileOut);
        }
    }
}
