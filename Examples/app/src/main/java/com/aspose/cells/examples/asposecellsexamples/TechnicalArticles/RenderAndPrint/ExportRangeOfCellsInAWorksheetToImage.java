package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExportRangeOfCellsInAWorksheetToImage {

    private static final String TAG = ExportRangeOfCellsInAWorksheetToImage.class.getName();

    public void exportRangeOfCellsInAWorksheetToImage() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook from source file.
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Access the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Set the print area with your desired range
            worksheet.getPageSetup().setPrintArea("E8:H15");

            //Set all margins as 0
            worksheet.getPageSetup().setLeftMargin(0);
            worksheet.getPageSetup().setRightMargin(0);
            worksheet.getPageSetup().setTopMargin(0);
            worksheet.getPageSetup().setBottomMargin(0);

            //Set OnePagePerSheet option as true
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.setOnePagePerSheet(true);
            options.setImageFormat(ImageFormat.getJpeg());

            //Take the image of worksheet
            SheetRender sr = new SheetRender(worksheet, options);
            sr.toImage(0, filePath + "ExportRangeOfCells_Out.jpg");
        } catch (Exception e) {
            Log.e(TAG, "Export Range of Cells in a Worksheet to Image", e);
        }
    }
}
