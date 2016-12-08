package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class ReleaseUnmanagedResourcesOfTheWorkbook {

    private static final String TAG = ReleaseUnmanagedResourcesOfTheWorkbook.class.getName();

    public void releaseUnmanagedResourcesOfTheWorkbook() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook();

            //Call dispose method
            //It performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
            workbook.dispose();

        } catch (Exception e) {
            Log.e(TAG, "Release Unmanaged Resources of the Workbook", e);
        }
    }

}
