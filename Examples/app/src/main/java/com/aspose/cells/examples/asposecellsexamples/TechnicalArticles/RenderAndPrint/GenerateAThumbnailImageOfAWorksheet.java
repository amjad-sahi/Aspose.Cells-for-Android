package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GenerateAThumbnailImageOfAWorksheet {
    private static final String TAG = GenerateAThumbnailImageOfAWorksheet.class.getName();

    public void generateAThumbnailImageOfAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate and open an Excel file
            Workbook book = new Workbook(filePath + "Book1.xlsx");

            //Define ImageOrPrintOptions
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            //Set the vertical and horizontal resolution
            imgOptions.setVerticalResolution(200);
            imgOptions.setHorizontalResolution(200);
            //Set the image's format
            imgOptions.setImageFormat(ImageFormat.getJpeg());
            //One page per sheet is enabled
            imgOptions.setOnePagePerSheet(true);

            //Get the first worksheet
            Worksheet sheet = book.getWorksheets().get(0);
            //Render the sheet with respect to specified image/print options
            SheetRender sr = new SheetRender(sheet, imgOptions);
            //Render the image for the sheet
            sr.toImage(0, filePath + "GenerateAThumbnailImage_Out.jpg");
        } catch (Exception e) {
            Log.e(TAG, "Generate a Thumbnail Image of a Worksheet", e);
        }
    }
}