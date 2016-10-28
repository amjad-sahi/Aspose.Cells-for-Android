package com.aspose.cells.examples.asposecellsexamples.DataFormatting.AdvanceFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ActivateSheetsAndMakeAnActiveCell {

    private static final String TAG = ActivateSheetsAndMakeAnActiveCell.class.getName();

    /**
     * Aspose.Cells provides some specific API for the tasks.
     * For example, the WorksheetCollection.setActiveSheet method is useful for setting the active sheet.
     * The Worksheet.setActiveCell method is used to set and get an active cell.
     */
    public void activateSheetAndMakeAnActiveCell() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the workbook.
            Worksheet worksheet1 = workbook.getWorksheets().get(0);

            //Get the cells in the worksheet.
            Cells cells = worksheet1.getCells();

            //Input data into B2 cell.
            cells.get("B2").setValue("Hello World!");

            //Set the first sheet as an active sheet.
            workbook.getWorksheets().setActiveSheetIndex(0);

            //Set B2 cell as an active cell in the worksheet.
            worksheet1.setActiveCell("B2");

            //Set the B column as the first visible column in the worksheet.
            worksheet1.setFirstVisibleColumn(1);

            //Set the 2nd row as the first visible row in the worksheet.
            worksheet1.setFirstVisibleRow(1);

            //Save the excel file.
            workbook.save(filePath + File.separator + "ActiveCell_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Activate Sheet and Make an Active Cell", e);
        }

    }
}
