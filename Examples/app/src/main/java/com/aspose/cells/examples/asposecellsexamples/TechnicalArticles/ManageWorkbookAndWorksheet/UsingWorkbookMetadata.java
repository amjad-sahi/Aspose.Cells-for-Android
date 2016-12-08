package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.MetadataOptions;
import com.aspose.cells.MetadataType;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookMetadata;

import java.io.File;

public class UsingWorkbookMetadata {

    private static final String TAG = UsingWorkbookMetadata.class.getName();

    public void usingWorkbookMetadata() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            // Open Workbook metadata
            MetadataOptions options = new MetadataOptions(MetadataType.DOCUMENT_PROPERTIES);
            WorkbookMetadata meta = new WorkbookMetadata(filePath + "sample.xlsx", options);

            // Set some properties
            meta.getCustomDocumentProperties().add("test", "testproperty");

            // Save the metadata info
            meta.save(filePath + "Sample2.xlsx");

            // Open the workbook
            Workbook w = new Workbook(filePath + "Sample2.xlsx");

            // Read document property
            Log.i(TAG, w.getCustomDocumentProperties().get("test").toString());
        } catch (Exception e) {
            Log.e(TAG, "Disable exporting frame scripts and document properties", e);
        }
    }
}