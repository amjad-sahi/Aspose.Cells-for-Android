package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;

import java.io.File;

public class SetPivotTableOptionForEmptyCellsShow {

    private static final String TAG = SetPivotTableOptionForEmptyCellsShow.class.getName();

    public void setPivotTableOptionForEmptyCellsShow() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook(filePath + "source.xlsx");

            PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);

            //Indicating if or not display the empty cell value
            pt.setDisplayNullString(true);

            //Indicating the null string
            pt.setNullString("null");

            pt.calculateData();

            pt.setRefreshDataOnOpeningFile(false);

            wb.save(filePath + "PivotTableOptionForEmptyCellsShow_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Setting Pivot Table Option - For Empty Cells Show", e);
        }
    }
}
