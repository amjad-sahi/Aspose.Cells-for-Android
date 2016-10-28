package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class ZoomFactor {

    private static final String TAG = ZoomFactor.class.getName();

    public void controllingTheZoomFactor() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            //Setting the zoom factor of the worksheet to 75
            worksheet.setZoom(75);

            workbook.save(filePath + File.separator + "ZoomFactor_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Controlling the Zoom Factor", e);
        }
    }
}