package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class DisableExportingFrameScriptsAndDocumentProperties {

    private static final String TAG = DisableExportingFrameScriptsAndDocumentProperties.class.getName();

    public void disableExportingFrameScriptsAndDocumentProperties() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            // Open the required workbook to convert
            Workbook workbook = new Workbook(filePath + "sample.xlsx");

            // Disable exporting frame scripts and document properties
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.setExportFrameScriptsAndProperties(false);

            // Save workbook as HTML
            workbook.save(filePath + "DisableExportingFrameScripts_Out.html", options);

            Log.i(TAG, "File saved");
        } catch (Exception e) {
            Log.e(TAG, "Disable exporting frame scripts and document properties", e);
        }
    }
}
