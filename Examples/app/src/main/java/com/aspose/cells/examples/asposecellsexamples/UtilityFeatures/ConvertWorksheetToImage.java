package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ConvertWorksheetToImage {

    private static final String TAG = ConvertWorksheetToImage.class.getName();

    public void convertToImage() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            //Set the image type
            imgOptions.setImageFormat(ImageFormat.getPng());
            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Create a SheetRender object for the target sheet
            SheetRender sr = new SheetRender(sheet, imgOptions);
            for (int j = 0; j < sr.getPageCount(); j++)  {
                //Generate an image for the worksheet
                sr.toImage(j, filePath + File.separator + "MySheetImage_" + j + "_Out.png");
            }

        } catch (Exception e) {
            Log.e(TAG, "Convert to Image", e);
        }
    }
}
