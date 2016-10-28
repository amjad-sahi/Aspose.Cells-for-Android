package com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup;

import com.aspose.cells.PageSetup;
import com.aspose.cells.PrintCommentsType;
import com.aspose.cells.PrintErrorsType;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class PrintOptions {
    private static final String TAG = PrintOptions.class.getName();

    public void setPrintArea() {
        //Instantiating a Workbook object
        Workbook workbook=new Workbook();

        //Accessing the first worksheet in the Workbook file
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);

        //Obtaining the reference of the PageSetup of the worksheet
        PageSetup pageSetup = sheet.getPageSetup();

        //Specifying the cells range (from A1 cell to T35 cell) of the print area
        pageSetup.setPrintArea("A1:T35");
    }

    public void setPrintTitles() {
        //Instantiating a Workbook object
        Workbook workbook=new Workbook();

        //Accessing the first worksheet in the Workbook file
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);

        //Obtaining the reference of the PageSetup of the worksheet
        PageSetup pageSetup = sheet.getPageSetup();

        //Defining column numbers A & B as title columns
        pageSetup.setPrintTitleColumns("$A:$B");

        //Defining row numbers 1 & 2 as title rows
        pageSetup.setPrintTitleRows("$1:$2");
    }

    public void setOtherPrintOptions() {
        //Instantiating a Workbook object
        Workbook workbook=new Workbook();

        //Accessing the first worksheet in the Workbook file
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);

        //Obtaining the reference of the PageSetup of the worksheet
        PageSetup pageSetup = sheet.getPageSetup();

        //Allowing to print gridlines
        pageSetup.setPrintGridlines(true);

        //Allowing to print row/column headings
        pageSetup.setPrintHeadings(true);

        //Allowing to print worksheet in black & white mode
        pageSetup.setBlackAndWhite (true);

        //Allowing to print comments as displayed on worksheet
        pageSetup.setPrintComments ( PrintCommentsType.PRINT_IN_PLACE);

        //Allowing to print worksheet with draft quality
        pageSetup.setPrintDraft(true);

        //Allowing to print cell errors as N/A
        pageSetup.setPrintErrors(PrintErrorsType.PRINT_ERRORS_NA);
    }

    public void setPageOrder() {
        //Instantiating a Workbook object
        Workbook workbook=new Workbook();

        //Accessing the first worksheet in the Workbook file
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);

        //Obtaining the reference of the PageSetup of the worksheet
        PageSetup pageSetup = sheet.getPageSetup();

        //Setting the printing order of the pages to over then down
        pageSetup.setOrder(PrintOrderType.OVER_THEN_DOWN);
    }

}
