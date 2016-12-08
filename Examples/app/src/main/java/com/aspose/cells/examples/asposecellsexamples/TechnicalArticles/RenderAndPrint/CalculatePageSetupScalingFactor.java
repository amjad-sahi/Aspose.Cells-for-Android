package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.util.Log;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class CalculatePageSetupScalingFactor {

    private static final String TAG = CalculatePageSetupScalingFactor.class.getName();

    public void calculatePageSetupScalingFactor() {
        try {
            //Create workbook object
            Workbook workbook = new Workbook();

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Put some data in these cells
            worksheet.getCells().get("A4").putValue("Test");
            worksheet.getCells().get("S4").putValue("Test");

            //Set paper size
            worksheet.getPageSetup().setPaperSize(PaperSizeType.PAPER_A_4);

            //Set fit to pages wide as 1
            worksheet.getPageSetup().setFitToPagesWide(1);

            //Calculate page scale via sheet render
            SheetRender sr = new SheetRender(worksheet, new ImageOrPrintOptions());

            //Write the page scale value
            Log.i(TAG, "Page Scale: " + sr.getPageScale());
        } catch (Exception e) {
            Log.e(TAG, "Calculate Page Setup Scaling Factor", e);
        }
    }

}
