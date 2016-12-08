package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ChangeLayoutOfPivotTable {

    private static final String TAG = ChangeLayoutOfPivotTable.class.getName();

    public void changeLayoutOfPivotTable() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first pivot table
            PivotTable pivotTable = worksheet.getPivotTables().get(0);

            //1 - Show the pivot table in compact form
            pivotTable.showInCompactForm();

            //Refresh the pivot table
            pivotTable.refreshData();
            pivotTable.calculateData();

            //Save the output
            workbook.save(filePath + "CompactForm_Out.xlsx");

            //2 - Show the pivot table in outline form
            pivotTable.showInOutlineForm();

            //Refresh the pivot table
            pivotTable.refreshData();
            pivotTable.calculateData();

            //Save the output
            workbook.save(filePath + "OutlineForm_Out.xlsx");

            //3 - Show the pivot table in tabular form
            pivotTable.showInTabularForm();

            //Refresh the pivot table
            pivotTable.refreshData();
            pivotTable.calculateData();

            //Save the output
            workbook.save(filePath + "TabularForm_Out.xlsx");

        } catch (Exception e) {
            Log.e(TAG, "Changing the Layout of Pivot Table", e);
        }
    }
}
