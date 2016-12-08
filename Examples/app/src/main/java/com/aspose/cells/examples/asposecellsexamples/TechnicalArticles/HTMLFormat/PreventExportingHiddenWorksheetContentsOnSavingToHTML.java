package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class PreventExportingHiddenWorksheetContentsOnSavingToHTML {
    private static final String TAG = PreventExportingHiddenWorksheetContentsOnSavingToHTML.class.getName();

    public void preventExportingHiddenWorksheetContentsOnSavingToHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Do not export hidden worksheet contents
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.setExportHiddenWorksheet(false);

            //Save the workbook
            workbook.save(filePath + "PreventExportingHiddenWorksheetContents_Out.html", options);
        } catch (Exception e) {
            Log.e(TAG, "Prevent Exporting Hidden Worksheet Contents on Saving to HTML", e);
        }
    }

}
