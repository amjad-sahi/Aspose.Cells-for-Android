package com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods;

import android.util.Log;

import com.aspose.cells.CellsHelper;

public class GetRowAndColumnIndicesFromCellName {

    private static final String TAG = GetRowAndColumnIndicesFromCellName.class.getName();

    public void getRowAndColumnIndicesFromCellName() {
        try {
            int[] cellIndices = CellsHelper.cellNameToIndex("C6");

            Log.v(TAG, "Row Index of Cell C6: " + cellIndices[0]);
            Log.v(TAG, "Column Index of Cell C6: " + cellIndices[1]);
        } catch (Exception e) {
            Log.e(TAG, "Get Row and Column Indices from Cell Name", e);
        }
    }
}
