package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class ConvertExcelFilesToHTML {

    private static final String TAG = ConvertExcelFilesToHTML.class.getName();

    public void convertExcelToHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Specify the HTML Saving Options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);

            //Save the HTML file
            workbook.save(filePath + File.separator + "ConvertExcelToHTML_Out.html", saveOptions);

        } catch (Exception e) {
            Log.e(TAG, "Convert Excel to HTML", e);
        }
    }

    public void setImagePreferencesForHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Create an instance of HtmlSaveOptions
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);

            //Set the ImageFormat to PNG
            saveOptions.getImageOptions().setImageFormat(ImageFormat.getPng());

            //Set the background of image as transparent
            saveOptions.getImageOptions().setTransparent(true);

            //Save spreadsheet to HTML while passing object of HtmlSaveOptions
            workbook.save(filePath + File.separator + "SetImagePreferencesForHTML_Out.html", saveOptions);

        } catch (Exception e) {
            Log.e(TAG, "Set Image Preferences for HTML", e);
        }
    }

}
