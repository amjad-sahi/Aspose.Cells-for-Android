package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExportWorksheetOrChartToImageWithDesiredDimensions {

    private static final String TAG = ExportWorksheetOrChartToImageWithDesiredDimensions.class.getName();

    public void exportWorksheetOrChartToImageWithDesiredDimensions() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook class and load an existing spreadsheet
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Access first Worksheet from the collection
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Create an instance of ImageOrPrintOptions class
            ImageOrPrintOptions opts = new ImageOrPrintOptions();
            //Set format for resultant image
            opts.setImageFormat(ImageFormat.getPng());
            //Set required dimensions of resultant image in unit of Pixels
            opts.setDesiredSize(400, 400);

            //Create an instance of SheetRender and initialize it with instances of Worksheet & ImageOrPrintOptions classes
            SheetRender sr = new SheetRender(worksheet, opts);
            //Convert the Worksheet to image
            sr.toImage(0, filePath + "ExportWorksheetOrChartToImage_Out.png");
        } catch (Exception e) {
            Log.e(TAG, "Export Worksheet or Chart to Image with Desired Dimensions", e);
        }
    }
}
