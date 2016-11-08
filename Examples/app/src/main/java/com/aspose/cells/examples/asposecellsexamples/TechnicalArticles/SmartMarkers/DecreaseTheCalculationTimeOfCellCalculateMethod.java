package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.SmartMarkers;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DecreaseTheCalculationTimeOfCellCalculateMethod {
    private static final String TAG = DecreaseTheCalculationTimeOfCellCalculateMethod.class.getName();

    public void decreaseTheCalculationTimeOfCellCalculateMethod() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Test calculation time after setting recursive true
            testCalcTimeRecursive(filePath, true);

            //Test calculation time after setting recursive false
            testCalcTimeRecursive(filePath, false);
        } catch (Exception e) {
            Log.e(TAG, "Decrease the Calculation Time of Cell.Calculate() method", e);
        }
    }

    private void testCalcTimeRecursive(String filePath, boolean rec) throws Exception {

        // Load your sample workbook
        Workbook wb = new Workbook(filePath + "CalculationTime.xlsx");

        // Access first worksheet
        Worksheet ws = wb.getWorksheets().get(0);

        // Set the calculation option, set recursive true or false as per parameter
        CalculationOptions opts = new CalculationOptions();
        opts.setRecursive(rec);

        // Start calculating time in nanoseconds
        long startTime = System.nanoTime();

        // Calculate cell A1 one million times
        for (int i = 0; i < 1000000; i++) {
            ws.getCells().get("A1").calculate(opts);
        }

        // Calculate elapsed time in seconds
        long second = 1000000000;
        long estimatedTime = System.nanoTime() - startTime;
        estimatedTime = estimatedTime / second;

        // Print the elapsed time in seconds
        Log.i(TAG, "Recursive " + rec + ": " + estimatedTime + " seconds");
    }
}
