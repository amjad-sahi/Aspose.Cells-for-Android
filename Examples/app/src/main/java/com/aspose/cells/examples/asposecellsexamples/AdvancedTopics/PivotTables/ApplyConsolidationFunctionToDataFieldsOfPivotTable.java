package com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ApplyConsolidationFunctionToDataFieldsOfPivotTable {

    private static final String TAG = ApplyConsolidationFunctionToDataFieldsOfPivotTable.class.getName();

    /**
     * This example applies Average consolidation function to first data field (or value field) and DistinctCount consolidation function to second data field (or value field).
     */
    public void applyConsolidationFunctionToDataFieldsOfPivotTable() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create workbook from source excel file
            Workbook workbook = new Workbook(filePath + File.separator + "source.xlsx");

            //Access the first worksheet of the workbook
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the first pivot table of the worksheet
            PivotTable pivotTable = worksheet.getPivotTables().get(0);

            //Apply Average consolidation function to first data field
            pivotTable.getDataFields().get(0).setFunction(ConsolidationFunction.AVERAGE);

            //Apply DistinctCount consolidation function to second data field
            pivotTable.getDataFields().get(1).setFunction(ConsolidationFunction.DISTINCT_COUNT);

            //Calculate the data to make changes affect
            pivotTable.calculateData();

            //Save the workbook
            workbook.save(filePath + File.separator + "ConsolidationFunction_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Apply ConsolidationFunction to Data Fields of Pivot Table", e);
        }
    }
}
