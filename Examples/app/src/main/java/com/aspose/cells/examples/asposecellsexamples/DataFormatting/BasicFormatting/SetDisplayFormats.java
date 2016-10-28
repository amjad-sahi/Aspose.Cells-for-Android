package com.aspose.cells.examples.asposecellsexamples.DataFormatting.BasicFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.util.Calendar;

public class SetDisplayFormats {

    private static final String TAG = SetDisplayFormats.class.getName();

    /**
     * Aspose.Cells offers some built-in number formats for configuring number and date display formats.
     * These built-in formats can be applied by using the Style object's setNumber method.
     * All built-in number formats are given unique numeric values.
     * Using the pre-defined number formats is faster than defining your own.
     */
    public void usingBuiltInNumberFormats() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue(Calendar.getInstance());

            //Setting the display format of the date to number 15 to show date as "d-mmm-yy"
            Style style = cell.getStyle();
            style.setNumber(15);
            cell.setStyle(style);

            //Adding a numeric value to "A2" cell
            cell = cells.get("A2");
            cell.setValue(20);

            //Setting the display format of the value to number 9 to show value as percentage
            style = cell.getStyle();
            style.setNumber(9);
            cell.setStyle(style);

            //Adding a numeric value to "A3" cell
            cell = cells.get("A3");
            cell.setValue(1546);

            //Setting the display format of the value to number 6 to show value as currency
            style = cell.getStyle();
            style.setNumber(6);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "UsingBuiltInNumberFormats_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Using BuiltIn Number Formats", e);
        }
    }

    /**
     * To define a customized format string and set a custom display format, use the Style object's setCustom method.
     * This approach is not as faster as using the in-built formats but is more flexible.
     */

    public void usingCustomNumberFormats() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue(Calendar.getInstance());

            //Setting the display format of the date to number 15 to show date as "d-mmm-yy"
            Style style = cell.getStyle();
            style.setCustom("d-mmm-yy");
            cell.setStyle(style);

            //Adding a numeric value to "A2" cell
            cell = cells.get("A2");
            cell.setValue(20);

            //Setting the display format of the value to number 9 to show value as percentage
            style = cell.getStyle();
            style.setCustom("0.0%");
            cell.setStyle(style);

            //Adding a numeric value to "A3" cell
            cell = cells.get("A3");
            cell.setValue(1546);

            //Setting the display format of the value to number 6 to show value as currency
            style = cell.getStyle();
            style.setCustom("$#,##0;[Red]$-#,##0");
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "UsingCustomNumberFormats_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Using Custom Number Formats", e);
        }
    }
}
