package com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods;

import android.util.Log;

import com.aspose.cells.CellsHelper;

public class GetCellNameFromRowAndColumnIndices {
    private static final String TAG = GetCellNameFromRowAndColumnIndices.class.getName();

    public void getCellNameFromRowAndColumnIndices() {
        try {
            String cellname = CellsHelper.cellIndexToName(0, 0);
            Log.v(TAG, "Cell Name at [0, 0]: " + cellname);

            cellname = CellsHelper.cellIndexToName(4, 0);
            Log.v(TAG, "Cell Name at [4, 0]: " + cellname);

            cellname = CellsHelper.cellIndexToName(0, 4);
            Log.v(TAG, "Cell Name at [0, 4]: " + cellname);

            cellname = CellsHelper.cellIndexToName(2, 2);
            Log.v(TAG, "Cell Name at [2, 2]: " + cellname);
        } catch (Exception e) {
            Log.e(TAG, "Get Cell Name from Row and Column Indices", e);
        }
    }

}
