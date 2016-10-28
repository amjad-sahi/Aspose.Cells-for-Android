package com.aspose.cells.examples.asposecellsexamples.Worksheets.SecurityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.Protection;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class AdvancedProtectionSettingsSinceExcelXP {
    private static final String TAG = AdvancedProtectionSettingsSinceExcelXP.class.getName();

    public void advanceProtectionSettings() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Protection protection = worksheet.getProtection();

            //Restricting users to delete columns of the worksheet
            protection.setAllowDeletingColumn(false);

            //Restricting users to delete row of the worksheet
            protection.setAllowDeletingRow(false);

            //Restricting users to edit contents of the worksheet
            protection.setAllowEditingContent(false);

            //Restricting users to edit objects of the worksheet
            protection.setAllowEditingObject(false);

            //Restricting users to edit scenarios of the worksheet
            protection.setAllowEditingScenario(false);

            //Restricting users to filter
            protection.setAllowFiltering(false);

            //Allowing users to format cells of the worksheet
            protection.setAllowFormattingCell(true);

            //Allowing users to format rows of the worksheet
            protection.setAllowFormattingRow(true);

            //Allowing users to insert columns in the worksheet
            protection.setAllowInsertingColumn(true);

            //Allowing users to insert hyperlinks in the worksheet
            protection.setAllowInsertingHyperlink(true);

            //Allowing users to insert rows in the worksheet
            protection.setAllowInsertingRow(true);

            //Allowing users to select locked cells of the worksheet
            protection.setAllowSelectingLockedCell(true);

            //Allowing users to select unlocked cells of the worksheet
            protection.setAllowSelectingUnlockedCell(true);

            //Allowing users to sort
            protection.setAllowSorting(true);

            //Allowing users to use pivot tables in the worksheet
            protection.setAllowUsingPivotTable(true);

            workbook.save(filePath + File.separator + "AdvanceProtectionSettings_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Advance Protection Settings", e);
        }
    }

    public void lockCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Cell cell = worksheet.getCells().get("A1");
            Style style = cell.getStyle();
            //Locking a cell
            style.setLocked(true);
            cell.setStyle(style);

        } catch(Exception e) {
            Log.e(TAG, "Lock Cells", e);
        }
    }

}
