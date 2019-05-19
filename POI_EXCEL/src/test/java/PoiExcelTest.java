
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
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
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
    XSSFWorkbook xsfwb;
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
        sheet = wb.createSheet("sheet1");
//        sheet2 = wb.createSheet("second sheet");
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
        Row row = sheet.createRow(0);
        // Create a cell and put a value in it.
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        // Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);
    }

//    private Row createRow() {
//        return sheet.createRow(0);
//    }
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
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        row.createCell(5).setCellType(CellType.ERROR);
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

    public void createFont() {

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);

        // Create a new font and alter it.
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 24);
        font.setFontName("Courier New");
        font.setItalic(true);
        font.setStrikeout(true);

        // Fonts are set into a style so create a new one to use.
        CellStyle style = wb.createCellStyle();
        style.setFont(font);

        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of fonts");
        cell.setCellStyle(style);

    }

    @Test
    public void customColor() throws FileNotFoundException, IOException {

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("custom XSSF colors");

        XSSFCellStyle style1 = wb.createCellStyle();
        style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap()));
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        try (OutputStream fileOut = new FileOutputStream(file_name)) {
            wb.write(fileOut);
        }
    }

    public void dataFormat() {

        sheet = wb.createSheet("format sheet");
        CellStyle style;
        DataFormat format = wb.createDataFormat();
        Row row;
        Cell cell;
        int rowNum = 0;
        int colNum = 0;

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("0.0"));
        cell.setCellStyle(style);

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("#,##0.0000"));
        cell.setCellStyle(style);

    }

    public void FitSheetOnePage() {

        sheet = wb.createSheet("format sheet");
        PrintSetup ps = sheet.getPrintSetup();
        sheet.setAutobreaks(true);
        ps.setFitHeight((short) 1);
        ps.setFitWidth((short) 1);

    }

    public void SetPageNumbersonFooter() {

        sheet = wb.createSheet("format sheet");
        Footer footer = sheet.getFooter();

//        footer.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());
    }

    public void repeatingRowAndColumn() {
        sheet = wb.createSheet("Sheet1");
        sheet2 = wb.createSheet("Sheet2");

        // Set the rows to repeat from row 4 to 5 on the first sheet.
        sheet.setRepeatingRows(CellRangeAddress.valueOf("4:5"));
        // Set the columns to repeat from column A to C on the second sheet
        sheet2.setRepeatingColumns(CellRangeAddress.valueOf("A:C"));
    }

    public void headerAndFooter() {

        sheet = wb.createSheet("new sheetFor");

        Header header = sheet.getHeader();
        header.setCenter("Center Header");
        header.setLeft("Left Header");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic")
                + HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");

    }

    @Test
    public void commentCell() throws FileNotFoundException, IOException {
        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

        CreationHelper factory = wb.getCreationHelper();

        sheet = wb.createSheet();

        Row row = sheet.createRow(3);
        Cell cell = row.createCell(5);
        cell.setCellValue("F4");

        Drawing drawing = sheet.createDrawingPatriarch();

        // When the comment box is visible, have it show in a 1x3 space
        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + 1);
        anchor.setRow1(row.getRowNum());
        anchor.setRow2(row.getRowNum() + 3);

        // Create the comment and set the text+author
        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = factory.createRichTextString("Hello, World!");
        comment.setString(str);
        comment.setAuthor("Apache POI");

        // Assign the comment to the cell
        cell.setCellComment(comment);

        String fname = "comment-xssf.xls";
        if (wb instanceof XSSFWorkbook) {
            fname += "x";
        }
        try (OutputStream fileOut = new FileOutputStream(file_name)) {
            wb.write(fileOut);
        }
    }

    @Test
    public void testColumnWidth() {

//        sheet = wb.createSheet();
        Row row = sheet.createRow(3);
        Cell cell = row.createCell(1);
        cell.setCellValue("F4asdasd sdasddasdasdasd");
        // Auto-size the columns.
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
    }

    public void createHiperlink() {
//        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        //cell style for hyperlinks
        //by default hyperlinks are blue and underlined
        CellStyle hlink_style = wb.createCellStyle();
        Font hlink_font = wb.createFont();
        hlink_font.setUnderline(Font.U_SINGLE);
        hlink_font.setColor(IndexedColors.BLUE.getIndex());
        hlink_style.setFont(hlink_font);

        Cell cell;
//        Sheet 
        sheet = wb.createSheet("Hyperlinks");
        //URL
        cell = sheet.createRow(0).createCell(0);
        cell.setCellValue("URL Link");

        Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress("http://poi.apache.org/");
        cell.setHyperlink(link);
        cell.setCellStyle(hlink_style);

        //link to a file in the current directory
        cell = sheet.createRow(1).createCell(0);
        cell.setCellValue("File Link");
        link = createHelper.createHyperlink(HyperlinkType.FILE);
        link.setAddress("link1.xls");
        cell.setHyperlink(link);
        cell.setCellStyle(hlink_style);

        //e-mail link
        cell = sheet.createRow(2).createCell(0);
        cell.setCellValue("Email Link");
        link = createHelper.createHyperlink(HyperlinkType.EMAIL);
        //note, if subject contains white spaces, make sure they are url-encoded
        link.setAddress("mailto:poi@apache.org?subject=Hyperlinks");
        cell.setHyperlink(link);
        cell.setCellStyle(hlink_style);

        //link to a place in this workbook
        //create a target sheet and cell
//        Sheet
        sheet2 = wb.createSheet("Target Sheet");
        sheet2.createRow(0).createCell(0).setCellValue("Target Cell");

        cell = sheet.createRow(3).createCell(0);
        cell.setCellValue("Worksheet Link");
        Hyperlink link2 = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        link2.setAddress("'Target Sheet'!A1");
        cell.setHyperlink(link2);
        cell.setCellStyle(hlink_style);
    }

    public void autofilter() {
//        Workbook wb = new HSSFWorkbook(); //or new XSSFWorkbook();
//        sheet = wb.createSheet();
        sheet.setAutoFilter(CellRangeAddress.valueOf("C5:F200"));
    }

    public void conditionalFormat() {

//        Workbook workbook = new HSSFWorkbook(); // or new XSSFWorkbook();
//        Sheet sheet = workbook.createSheet();
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "0");
        FontFormatting fontFmt = rule1.createFontFormatting();
        fontFmt.setFontStyle(true, false);
        fontFmt.setFontColorIndex(IndexedColors.DARK_RED.index);

        BorderFormatting bordFmt = rule1.createBorderFormatting();
        bordFmt.setBorderBottom(BorderStyle.THIN);
        bordFmt.setBorderTop(BorderStyle.THICK);
        bordFmt.setBorderLeft(BorderStyle.DASHED);
        bordFmt.setBorderRight(BorderStyle.DOTTED);

        PatternFormatting patternFmt = rule1.createPatternFormatting();
        patternFmt.setFillBackgroundColor(IndexedColors.YELLOW.index);

        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.BETWEEN, "-10", "10");
        ConditionalFormattingRule[] cfRules
                = {
                    rule1, rule2
                };

        CellRangeAddress[] regions = {
            CellRangeAddress.valueOf("A3:A5")
        };

        sheetCF.addConditionalFormatting(regions, cfRules);
    }

    public void outline() {
//        Workbook wb = new HSSFWorkbook();
//        Sheet sheet1 = wb.createSheet("new sheet");
        sheet.groupRow(5, 14);
        sheet.groupRow(7, 14);
        sheet.groupRow(16, 19);
        sheet.groupColumn(4, 7);
        sheet.groupColumn(9, 12);
        sheet.groupColumn(10, 11);
    }

    public void settingCellProperties() {
//        Workbook workbook = new XSSFWorkbook();  // OR new HSSFWorkbook()
//        Sheet sheet = workbook.createSheet("Sheet1");
        Map<String, Object> properties = new HashMap<>();

        // border around a cell
        properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);

        // Give it a color (RED)
        properties.put(CellUtil.TOP_BORDER_COLOR, IndexedColors.RED.getIndex());
        properties.put(CellUtil.BOTTOM_BORDER_COLOR, IndexedColors.RED.getIndex());
        properties.put(CellUtil.LEFT_BORDER_COLOR, IndexedColors.RED.getIndex());
        properties.put(CellUtil.RIGHT_BORDER_COLOR, IndexedColors.RED.getIndex());

        // Apply the borders to the cell at B2
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        CellUtil.setCellStyleProperties(cell, properties);

        // Apply the borders to a 3x3 region starting at D4
        for (int ix = 3; ix <= 5; ix++) {
            row = sheet.createRow(ix);
            for (int iy = 3; iy <= 5; iy++) {
                cell = row.createCell(iy);
                CellUtil.setCellStyleProperties(cell, properties);
            }
        }
    }

    public void pivotTable() {
//        XSSFWorkbook wb = new XSSFWorkbook();
//        XSSFSheet sheet = wb.createSheet();

        //Create some data to build the pivot table on
//        setCellData(sheet);
//        XSSFPivotTable pivotTable = sheet.createPivotTable(new AreaReference("A1:D4"), new CellReference("H5"));
        //Configure the pivot table
        //Use first column as row label
//        pivotTable.addRowLabel(0);
//        //Sum up the second column
//        pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 1);
//        //Set the third column as filter
//        pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, 2);
//        //Add filter on forth column
//        pivotTable.addReportFilter(3);
    }

    public void multiplyStyle() {
        // XSSF Example
        Row row = sheet.createRow(3);
        XSSFCell cell = (XSSFCell) row.createCell(1);
        XSSFRichTextString rt = new XSSFRichTextString("The quick brown fox");

        XSSFFont font1 = (XSSFFont) wb.createFont();
        font1.setBold(true);
        font1.setColor(new XSSFColor(new java.awt.Color(255, 0, 0)));
        rt.applyFont(0, 10, font1);

        XSSFFont font2 = (XSSFFont) wb.createFont();
        font2.setItalic(true);
        font2.setUnderline(XSSFFont.U_DOUBLE);
        font2.setColor(new XSSFColor(new java.awt.Color(0, 255, 0)));
        rt.applyFont(10, 19, font2);

        XSSFFont font3 = (XSSFFont) wb.createFont();
        font3.setColor(new XSSFColor(new java.awt.Color(0, 0, 255)));
        rt.append(" Jumped over the lazy dog", font3);

        cell.setCellValue(rt);
    }

    @Test
    public void testWorkBook() throws FileNotFoundException, IOException {
//        testCell();
//        createDateCell();
//        differentKindOfCell();
//        fillColors();
//        mergCell();
//        createFont();
//        customColor();    
//        dataFormat();
//        FitSheetOnePage();
//        repeatingRowAndColumn();
//        headerAndFooter();
//        commentCell();
//        testColumnWidth();
//        createHiperlink();
//        autofilter();
//        conditionalFormat();
//        outline();
//        settingCellProperties();
        multiplyStyle();
        try (OutputStream fileOut = new FileOutputStream(file_name)) {
            wb.write(fileOut);
        }
    }
}
