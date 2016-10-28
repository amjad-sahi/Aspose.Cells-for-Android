package com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.SmartMarkersAndFormulaCalculation;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;

import java.io.File;

public class SmartMarkers {

    private static final String TAG = SmartMarkers.class.getName();

    public void smartMarkers() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a WorkbookDesigner object
            WorkbookDesigner designer = new WorkbookDesigner();

            //Set workbook which containing smart markers
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");
            designer.setWorkbook(workbook);

            //Set the data source for the designer spreadsheet
            //designer.setDataSource(dataSet);

            //Process the smart markers
            designer.process();

        } catch (Exception e) {
            Log.e(TAG, "Smart Markers", e);
        }
    }
}
