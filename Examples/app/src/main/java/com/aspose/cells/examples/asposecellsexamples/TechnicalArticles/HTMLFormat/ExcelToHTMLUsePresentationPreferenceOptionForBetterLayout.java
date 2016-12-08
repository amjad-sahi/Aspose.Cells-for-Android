package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class ExcelToHTMLUsePresentationPreferenceOptionForBetterLayout {

    private static final String TAG = ExcelToHTMLUsePresentationPreferenceOptionForBetterLayout.class.getName();

    public void excelToHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate the Workbook
            //Load an Excel file
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Create HtmlSaveOptions object
            HtmlSaveOptions options = new HtmlSaveOptions();

            //Set the Presenation preference option
            options.setPresentationPreference(true);

            //Save the Excel file to HTML with specified option
            workbook.save(filePath + "ExcelToHTML_Out.html", options);
        } catch (Exception e) {
            Log.e(TAG, "Excel to HTML", e);
        }
    }

}
