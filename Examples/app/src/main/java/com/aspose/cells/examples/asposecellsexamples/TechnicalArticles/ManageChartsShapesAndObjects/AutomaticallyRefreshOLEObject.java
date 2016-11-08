package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AutomaticallyRefreshOLEObject {

    private static final String TAG = AutomaticallyRefreshOLEObject.class.getName();

    public void automaticallyRefreshOLEObject() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook(filePath + "sample-oleobject.xlsx");

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Set auto load property of first Ole Object to true
            sheet.getOleObjects().get(0).setAutoLoad(true);

            //Save the result in XLSX format
            book.save(filePath + "AutomaticallyRefreshOLEObject_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Automatically Refresh OLE Object", e);
        }
    }
}
