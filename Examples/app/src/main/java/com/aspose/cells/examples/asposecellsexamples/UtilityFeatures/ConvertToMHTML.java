package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class ConvertToMHTML {

    private static final String TAG = ConvertToMHTML.class.getName();

    public void convertToMHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Specify the HTML saving options
            HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions(SaveFormat.M_HTML);

            //Save the MHT file
            workbook.save(filePath + File.separator + "ConvertToMHTML_Out.mht", htmlSaveOptions);
        } catch (Exception e) {
            Log.e(TAG, "Convert to MHTML", e);
        }
    }

    public void ConvertFromMHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Source.html");
            workbook.save(filePath + File.separator + "ConvertFromMHTML_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Convert From MHTML", e);
        }
    }
}
