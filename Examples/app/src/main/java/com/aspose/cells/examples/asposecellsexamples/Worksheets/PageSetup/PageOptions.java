package com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup;

import com.aspose.cells.PageOrientationType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class PageOptions {

    private static final String TAG = PageOptions.class.getName();

    public void pageOrientation() {
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        int sheetIndex = worksheets.add();
        Worksheet sheet = worksheets.get(sheetIndex);

        //Setting the orientation to Portrait
        PageSetup pageSetup = sheet.getPageSetup();
        pageSetup.setOrientation(PageOrientationType.PORTRAIT);
    }

    public void scalingFactor() {
        //Instantiating a Excel object
        Workbook workbook = new Workbook();

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        int sheetIndex = worksheets.add();
        Worksheet sheet = worksheets.get(sheetIndex);

        //Setting the scaling factor to 100
        PageSetup pageSetup = sheet.getPageSetup();
        pageSetup.setZoom(100);
    }

    public void fitToPageOptions() {
        // Instantiating a Workbook object
        Workbook workbook = new Workbook();

        // Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        int sheetIndex = worksheets.add();
        Worksheet sheet = worksheets.get(sheetIndex);

        PageSetup pageSetup = sheet.getPageSetup();

        // Setting the number of pages to which the length of the worksheet will
        // be spanned
        pageSetup.setFitToPagesTall(1);

        // Setting the number of pages to which the width of the worksheet will
        // be spanned
        pageSetup.setFitToPagesWide(1);
    }

    public void paperSize() {
        //Instantiating a Workbook object
        Workbook workbook =new Workbook();

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        int sheetIndex = worksheets.add();
        Worksheet sheet = worksheets.get(sheetIndex);

        //Setting the paper size to A4
        PageSetup pageSetup = sheet.getPageSetup();
        pageSetup.setPaperSize(PaperSizeType.PAPER_A_4);
    }

    public void printQuality() {
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        int sheetIndex = worksheets.add();
        Worksheet sheet = worksheets.get(sheetIndex);

        //Setting the print quality of the worksheet to 180 dpi
        PageSetup pageSetup = sheet.getPageSetup();
        pageSetup.setPrintQuality(180);
    }

    public void firstPageNumber() {
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();
        int sheetIndex = worksheets.add();
        Worksheet sheet = worksheets.get(sheetIndex);

        //Setting the first page number of the worksheet pages
        PageSetup pageSetup = sheet.getPageSetup();
        pageSetup.setFirstPageNumber(2);
    }
}
