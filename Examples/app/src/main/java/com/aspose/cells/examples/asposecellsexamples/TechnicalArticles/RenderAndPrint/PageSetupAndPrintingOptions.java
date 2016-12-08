package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PageOrientationType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.PrintCommentsType;
import com.aspose.cells.PrintErrorsType;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class PageSetupAndPrintingOptions {

    private static final String TAG = PageSetupAndPrintingOptions.class.getName();

    public void pageSetupOptions() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + "CustomerReport.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet sheet = workbook.getWorksheets().get(0);

            PageSetup pageSetup = sheet.getPageSetup();

            //Setting the orientation to Portrait
            pageSetup.setOrientation(PageOrientationType.PORTRAIT);

            //Setting the scaling factor to 100
            //pageSetup.setZoom(100);
            //OR Alternately you can use Fit to Page Options as under

            //Setting the number of pages to which the length of the worksheet will be spanned
            pageSetup.setFitToPagesTall(1);

            //Setting the number of pages to which the width of the worksheet will be spanned
            pageSetup.setFitToPagesWide(1);

            //Setting the paper size to A4
            pageSetup.setPaperSize(PaperSizeType.PAPER_A_4);

            //Setting the print quality of the worksheet to 1200 dpi
            pageSetup.setPrintQuality(1200);

            //Setting the first page number of the worksheet pages
            pageSetup.setFirstPageNumber(2);

            //Save the workbook
            workbook.save(filePath + "PageSetupOptions_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Page Setup Options", e);
        }
    }

    public void printOptions() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + "PageSetup.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet sheet = workbook.getWorksheets().get(0);

            PageSetup pageSetup = sheet.getPageSetup();

            //Specifying the cells range (from A1 cell to E30 cell) of the print area
            pageSetup.setPrintArea("A1:E30");

            //Defining column numbers A & E as title columns
            pageSetup.setPrintTitleColumns("$A:$E");

            //Defining row numbers 1 & 2 as title rows
            pageSetup.setPrintTitleRows("$1:$2");

            //Allowing to print gridlines
            pageSetup.setPrintGridlines(true);

            //Allowing to print row/column headings
            pageSetup.setPrintHeadings(true);

            //Allowing to print worksheet in black & white mode
            pageSetup.setBlackAndWhite(true);

            //Allowing to print comments as displayed on worksheet
            pageSetup.setPrintComments(PrintCommentsType.PRINT_IN_PLACE);

            //Allowing to print worksheet with draft quality
            pageSetup.setPrintDraft(true);

            //Allowing to print cell errors as N/A
            pageSetup.setPrintErrors(PrintErrorsType.PRINT_ERRORS_NA);

            //Setting the printing order of the pages to over then down
            pageSetup.setOrder(PrintOrderType.OVER_THEN_DOWN);

            //Save the workbook
            workbook.save(filePath + "PrintOptions_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Set Print options", e);
        }
    }

}
