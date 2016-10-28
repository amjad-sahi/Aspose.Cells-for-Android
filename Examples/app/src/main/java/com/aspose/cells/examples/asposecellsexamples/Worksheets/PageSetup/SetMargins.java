package com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class SetMargins {

    private static final String TAG = SetMargins.class.getName();

    public void pageMargins() {
        //Create a workbook object
        Workbook workbook = new Workbook();

        //Get the worksheets in the workbook
        WorksheetCollection worksheets = workbook.getWorksheets();

        //Get the first (default) worksheet
        Worksheet worksheet = worksheets.get(0);

        //Get the pagesetup object
        PageSetup pageSetup = worksheet.getPageSetup();

        //Set bottom,left,right and top page margins
        pageSetup.setBottomMargin(2);
        pageSetup.setLeftMargin(1);
        pageSetup.setRightMargin(1);
        pageSetup.setTopMargin(3);
    }

    public void centerOnPage() {
        //Create a workbook object
        Workbook workbook = new Workbook();

        //Get the worksheets in the workbook
        WorksheetCollection worksheets = workbook.getWorksheets();

        //Get the first (default) worksheet
        Worksheet worksheet = worksheets.get(0);

        //Get the pagesetup object
        PageSetup pageSetup = worksheet.getPageSetup();

        //Specify Center on page Horizontally and Vertically
        pageSetup.setCenterHorizontally(true);
        pageSetup.setCenterVertically(true);
    }

    public void headerAndFooterMargins() {
        //Create a workbook object
        Workbook workbook = new Workbook();

        //Get the worksheets in the workbook
        WorksheetCollection worksheets = workbook.getWorksheets();

        //Get the first (default) worksheet
        Worksheet worksheet = worksheets.get(0);

        //Get the pagesetup object
        PageSetup pageSetup = worksheet.getPageSetup();

        //Specify Header / Footer margins
        pageSetup.setHeaderMargin(2);
        pageSetup.setFooterMargin(2);
    }
}
