package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ConvertWorksheetToSVG {

    private static final String TAG = ConvertWorksheetToSVG.class.getName();

    public void convertWorksheetToSVG() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Create an instance of ImageOrPrintOptions & set SaveFormat as SVG
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            imgOptions.setSaveFormat(SaveFormat.SVG);

            //Set OnePagePerSheet to true in order to render all contents on one image
            imgOptions.setOnePagePerSheet(true);

            //Get count of all Worksheets to loop over them
            int sheetCount = workbook.getWorksheets().getCount();
            for(int i=0; i<sheetCount; i++) {
                Worksheet sheet = workbook.getWorksheets().get(i);
                //Create and initialize an instance of SheetRender
                SheetRender sr = new SheetRender(sheet, imgOptions);
                //Iterate over the pages
                for (int k = 0; k < sr.getPageCount(); k++) {
                    //Output the worksheet into Svg image format
                    sr.toImage(k, filePath + File.separator + sheet.getName() + k + "_Out.svg");
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Convert to Image", e);
        }
    }
}
