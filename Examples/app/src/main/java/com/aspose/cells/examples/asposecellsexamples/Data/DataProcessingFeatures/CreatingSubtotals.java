package com.aspose.cells.examples.asposecellsexamples.Data.DataProcessingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.Cells;
import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.Workbook;

import java.io.File;

public class CreatingSubtotals {

    private static final String TAG = CreatingSubtotals.class.getName();

    public void creatingSubtotals() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new workbook
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Get the Cells collection in the first worksheet
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Create a cellarea i.e.., B3:C19
            CellArea ca = new CellArea();
            ca.StartRow = 2;
            ca.StartColumn = 1;
            ca.EndRow = 18;
            ca.EndColumn = 2;

            //Apply subtotal, the consolidation function is Sum and it will applied to
            //Second column (C) in the list
            cells.subtotal(ca, 0, ConsolidationFunction.SUM, new int[]{1});

            //Save the Excel file
            workbook.save(filePath + File.separator + "CreatingSubtotals_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Creating Subtotals", e);
        }
    }
}
