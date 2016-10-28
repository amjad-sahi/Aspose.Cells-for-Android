package com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ChangeAPivotTableSourceData {
    private static final String TAG = ChangeAPivotTableSourceData.class.getName();

    public void changeAPivotTableSourceData() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "pivot.xlsm");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Populating new data to the worksheet cells
            Cells cells = worksheet.getCells();
            Cell cell = cells.get("A9");
            cell.setValue("Golf");
            cell = cells.get("B9");
            cell.setValue("Qtr4");
            cell = cells.get("C9");
            cell.setValue(7000);

            //Changing named range "DataSource"
            Range range = cells.createRange(0, 0, 8, 2);
            range.setName("DataSource");

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "PivotTable_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Pivot Table", e);
        }
    }
}
