package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.io.FileInputStream;

public class SetBackgroundPictureForAWorksheet {
    private static final String TAG = SetBackgroundPictureForAWorksheet.class.getName();

    public void setAWorksheetBackground() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();
            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Get the image file.
            File file = new File(filePath + "school.jpg");
            //Get the picture into the streams.
            byte[] imageData = new byte[(int)file.length()];
            FileInputStream fis = new FileInputStream(file);
            fis.read(imageData);

            //Set the background image for the sheet.
            sheet.setBackgroundImage(imageData);

            //Save the Excel file
            workbook.save(filePath + "BackImageSheet_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set a Worksheet Background", e);
        }
    }
}
