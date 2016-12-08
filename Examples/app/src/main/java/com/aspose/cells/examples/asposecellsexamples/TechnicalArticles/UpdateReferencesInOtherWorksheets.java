package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.DeleteOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class UpdateReferencesInOtherWorksheets {

    private static final String TAG = UpdateReferencesInOtherWorksheets.class.getName();

    /**
     * Update references in other worksheets while deleting blank columns and rows in a worksheet
     */
    public void updateReferencesInOtherWorksheetsWhileDeletingBlankColumnsAndRowsInAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook
            Workbook wb = new Workbook();

            //Add second sheet with name Sheet2
            wb.getWorksheets().add("Sheet2");

            //Access first sheet and add some integer value in cell C1
            //Also add some value in any cell to increase the number of blank rows
            //and columns
            Worksheet sht1 = wb.getWorksheets().get(0);
            sht1.getCells().get("C1").putValue(4);
            sht1.getCells().get("K30").putValue(4);

            //Access second sheet and add formula in cell E3 which refers to cell
            //C1 in first sheet
            Worksheet sht2 = wb.getWorksheets().get(1);
            sht2.getCells().get("E3").setFormula("'Sheet1'!C1");

            //Calculate formulas of workbook
            wb.calculateFormula();

            //Print the formula and value of cell E3 in second sheet before
            //deleting blank columns and rows in Sheet1.
            Log.i(TAG, "Cell E3 before deleting blank columns and rows in Sheet1.");
            Log.i(TAG, "--------------------------------------------------------");
            Log.i(TAG, "Cell Formula: " + sht2.getCells().get("E3").getFormula());
            Log.i(TAG, "Cell Value: " + sht2.getCells().get("E3").getStringValue());

            //If you comment DeleteOptions.UpdateReference property below, then the
            //formula in cell E3 in second sheet will not be updated
            DeleteOptions opts = new DeleteOptions();
            opts.setUpdateReference(true);

            //Delete all blank rows and columns with delete options
            sht1.getCells().deleteBlankColumns(opts);
            sht1.getCells().deleteBlankRows(opts);

            //Calculate formulas of workbook
            wb.calculateFormula();

            //Print the formula and value of cell E3 in second sheet after deleting
            //blank columns and rows in Sheet1.
            Log.i(TAG, "");
            Log.i(TAG, "");
            Log.i(TAG, "Cell E3 after deleting blank columns and rows in Sheet1.");
            Log.i(TAG, "--------------------------------------------------------");
            Log.i(TAG, "Cell Formula: " + sht2.getCells().get("E3").getFormula());
            Log.i(TAG, "Cell Value: " + sht2.getCells().get("E3").getStringValue());
        } catch (Exception e) {
            Log.e(TAG, "Update references in other worksheets while deleting blank columns and rows in a worksheet", e);
        }
    }
}
