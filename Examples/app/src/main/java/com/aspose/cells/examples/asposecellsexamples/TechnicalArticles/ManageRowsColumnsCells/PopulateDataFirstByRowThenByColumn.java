package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;

public class PopulateDataFirstByRowThenByColumn {
    private static final String TAG = PopulateDataFirstByRowThenByColumn.class.getName();

    public void populateDataFirstByRowThenByColumn() {
        try {
            Workbook workbook = new Workbook();
            Cells cells = workbook.getWorksheets().get(0).getCells();
            cells.get("A1").setValue("data1");
            cells.get("B1").setValue("data2");
            cells.get("A2").setValue("data3");
            cells.get("B2").setValue("data4");
        } catch (Exception e) {
            Log.e(TAG, "Populate Data First by Row then by Column", e);
        }
    }
}
