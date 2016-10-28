package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class FreezePanes {
    private static final String TAG = FreezePanes.class.getName();

    public void setFreezePanes() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            //Applying freeze panes settings
            worksheet.freezePanes(3,2,3,2);

            workbook.save(filePath + File.separator + "FreezePanes_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Freeze Panes", e);
        }
    }

}
