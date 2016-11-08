package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PivotTable;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class RefreshAndCalculatePivotTableHavingCalculatedItems {

    private static final String TAG = RefreshAndCalculatePivotTableHavingCalculatedItems.class.getName();

    public void refreshAndCalculatePivotTableHavingCalculatedItems() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load source spreadsheet containing a pivot table having calculated items
            Workbook book = new Workbook(filePath + "sample-pivottable.xlsx");

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Change the value of cell D2
            sheet.getCells().get("D2").putValue(20);

            //Refresh and calculate all the pivot tables inside the worksheet
            for (int i = 0; i < sheet.getPivotTables().getCount(); i++) {
                PivotTable pivot = sheet.getPivotTables().get(i);
                pivot.refreshData();
                pivot.calculateData();
            }

            //Save the result in PDF format
            book.save(filePath + "RefreshAndCalculatePivotTable_Out.pdf", SaveFormat.PDF);
        } catch (Exception e) {
            Log.e(TAG, "Refresh and Calculate Pivot Table having Calculated Items", e);
        }
    }
}
