package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DetectIfWorksheetIsPasswordProtected {

    private static final String TAG = DetectIfWorksheetIsPasswordProtected.class.getName();

    public void detectIfWorksheetIsPasswordProtected() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load a spreadsheet
            Workbook book = new Workbook(filePath + "sample.xlsx");

            //Access the protected Worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Check if Worksheet is password protected
            if (sheet.getProtection().isProtectedWithPassword()) {
                Log.i(TAG, "Worksheet is password protected");
            }
        } catch (Exception e) {
            Log.e(TAG, "Detect if Worksheet is Password Protected", e);
        }
    }

}
