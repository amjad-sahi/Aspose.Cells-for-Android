package com.aspose.cells.examples.asposecellsexamples.RowsAndColumns;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CopyRowsAndColumns {

    private static final String TAG = CopyRowsAndColumns.class.getName();

    public void copyRows() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Copy the second row with data, formatting, images and drawing objects
            //to the 12th row in the worksheet.
            worksheet.getCells().copyRow(worksheet.getCells(), 1, 11);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "CopyRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Copy Rows", e);
        }
    }

    public void copyColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Create a new Workbook.
            Workbook excelWorkbook0 = new Workbook();

            // Get the first worksheet in the book.
            Worksheet ws0 = excelWorkbook0.getWorksheets().get(0);

            // Put some data into header rows (A1:A4)
            for (int i = 1; i < 5; i++) {
                ws0.getCells().get("A" + i).setValue("Header Row " + i);
            }

            // Put some detail data (A5:A999)
            for (int i = 5; i < 1000; i++) {
                ws0.getCells().get("A" + i).setValue("Detail Row " + i);
            }

            // Create another Workbook.
            Workbook excelWorkbook1 = new Workbook();

            // Get the first worksheet in the book.
            Worksheet ws1 = excelWorkbook1.getWorksheets().get(0);

            // Copy the first column from the first worksheet of the first workbook into
            // the first worksheet of the second workbook.
            ws1.getCells().copyColumn(ws0.getCells(), 0, 2);

            // Autofit the column.
            ws1.autoFitColumn(2);

            // Save the Excel file.
            excelWorkbook1.save(filePath + File.separator + "CopyColumn.xls");
        } catch (Exception e) {
            Log.e(TAG, "Copy Columns", e);
        }
    }

}
