package com.aspose.cells.examples.asposecellsexamples.Worksheets.SecurityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Protection;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.Style;

import java.io.File;

public class ProtectWorksheet {

    private static final String TAG = ProtectWorksheet.class.getName();

    public void protectAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Protection protection = worksheet.getProtection();

            //The following 3 methods are only for Excel 2000 and earlier formats
            protection.setAllowEditingContent(false);
            protection.setAllowEditingObject(false);
            protection.setAllowEditingScenario(false);

            //Protects the first worksheet with a password "1234"
            protection.setPassword("1234");

            workbook.save(filePath + File.separator + "ProtectAWorksheet_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Protect a Worksheet", e);
        }
    }

    public void protectCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Create a new workbook.
            Workbook workbook = new Workbook();

            //Accessing the first worksheet in the Excel file
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            // Define the style object.
            Style style;

            // Loop through all the columns in the worksheet and unlock them.
            for(int i = 0; i <= 255; i ++) {
                style = worksheet.getCells().getColumns().get(i).getStyle();
                style.setLocked(false);
                worksheet.getCells().getColumns().get(i).applyStyle(style, new StyleFlag());
            }

            // Lock the three cells...i.e. A1, B1, C1.
            style = worksheet.getCells().get("A1").getStyle();
            style.setLocked(true);
            worksheet.getCells().get("A1").setStyle(style);
            style = worksheet.getCells().get("B1").getStyle();
            style.setLocked(true);
            worksheet.getCells().get("B1").setStyle(style);
            style = worksheet.getCells().get("C1").getStyle();
            style.setLocked(true);
            worksheet.getCells().get("C1").setStyle(style);

            workbook.save(filePath + File.separator + "ProtectCells_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Protect Cells", e);
        }
    }

    public void protectARow() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Create a new workbook.
            Workbook workbook = new Workbook();

            //Accessing the first worksheet in the Excel file
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            // Define the style object.
            Style style;

            // Define the styleflag object.
            StyleFlag flag;

            // Loop through all the columns in the worksheet and unlock them.
            for(int i = 0; i <= 255; i ++) {
                style = worksheet.getCells().getColumns().get(i).getStyle();
                style.setLocked(false);
                flag = new StyleFlag();
                flag.setLocked(true);
                worksheet.getCells().getColumns().get(i).applyStyle(style, flag);
            }

            // Get the first row style.
            style = worksheet.getCells().getRows().get(0).getStyle();

            // Lock it.
            style.setLocked(true);

            // Instantiate the flag.
            flag = new StyleFlag();

            // Set the lock setting.
            flag.setLocked(true);

            // Apply the style to the first row.
            worksheet.getCells().getRows().get(0).applyStyle(style, flag);

            workbook.save(filePath + File.separator + "ProtectARow_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Protect A Row", e);
        }
    }

    public void protectAColumn() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Create a new workbook.
            Workbook workbook = new Workbook();

            //Accessing the first worksheet in the Excel file
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            // Define the style object.
            Style style;

            // Define the styleflag object.
            StyleFlag flag;

            // Loop through all the columns in the worksheet and unlock them.
            for(int i = 0; i <= 255; i ++) {
                style = worksheet.getCells().getColumns().get(i).getStyle();
                style.setLocked(false);
                flag = new StyleFlag();
                flag.setLocked(true);
                worksheet.getCells().getColumns().get(i).applyStyle(style, flag);
            }

            // Get the first column style.
            style = worksheet.getCells().getColumns().get(0).getStyle();

            // Lock it.
            style.setLocked(true);

            // Instantiate the flag.
            flag = new StyleFlag();

            // Set the lock setting.
            flag.setLocked(true);

            // Apply the style to the first column.
            worksheet.getCells().getColumns().get(0).applyStyle(style, flag);

            workbook.save(filePath + File.separator + "ProtectAColumn_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Protect A Column", e);
        }
    }
}








