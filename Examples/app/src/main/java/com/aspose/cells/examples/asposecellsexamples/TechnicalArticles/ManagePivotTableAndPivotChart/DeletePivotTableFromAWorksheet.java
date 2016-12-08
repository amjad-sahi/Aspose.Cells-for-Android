package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DeletePivotTableFromAWorksheet {

    private static final String TAG = DeletePivotTableFromAWorksheet.class.getName();

    public void deletePivotTableFromAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source Excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the first pivot table object
            PivotTable pivotTable = worksheet.getPivotTables().get(0);

            //Remove pivot table using pivot table object
            worksheet.getPivotTables().remove(pivotTable);

            //Remove pivot table using pivot table position
            worksheet.getPivotTables().removeAt(0);

            //Save the workbook
            workbook.save(filePath + "DeletePivotTable_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Delete Pivot Table from a Worksheet", e);
        }
    }
}
