package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExtractImageFromWorksheet {

    private static final String TAG = ExtractImageFromWorksheet.class.getName();

    public void extractImageFromWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load an existing spreadsheet
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Access the first worksheet from the collection
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the first picture from the collection of pictures
            Picture pic = worksheet.getPictures().get(0);

            //Define ImageOrPrintOptions
            ImageOrPrintOptions printoption = new ImageOrPrintOptions();

            //Specify the image format
            printoption.setImageFormat(ImageFormat.getJpeg());

            String fileName = "picture";

            //Save the image
            pic.toImage(filePath + fileName + ".jpg" , printoption);
        } catch (Exception e) {
            Log.e(TAG, "Extract Image from Worksheet", e);
        }
    }
}
