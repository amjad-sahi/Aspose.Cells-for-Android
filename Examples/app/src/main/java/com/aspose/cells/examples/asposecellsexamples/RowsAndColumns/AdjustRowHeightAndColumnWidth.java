package com.aspose.cells.examples.asposecellsexamples.RowsAndColumns;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AdjustRowHeightAndColumnWidth {

    private static final String TAG = AdjustRowHeightAndColumnWidth.class.getName();

    public void setRowHeight() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Setting the height of the second row to 13
            cells.setRowHeight(1, 13);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "SetRowHeight_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Row Height", e);
        }
    }

    public void setRowHeightForAllRows() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Set the height of all rows in the worksheet to 15
            worksheet.getCells().setStandardHeight(15f);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "SetRowHeightForAllRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Row Height For All Rows", e);
        }
    }

    public void setColumnWidth() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Set the width of the second column to 17.5
            cells.setColumnWidth(1, 17.5);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "SetColumnWidth_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Column Width", e);
        }
    }

    public void setWidthOfAllColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Set the width of all columns in the worksheet to 20.5
            worksheet.getCells().setStandardWidth(20.5f);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "SetWidthOfAllColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Width Of All Columns", e);
        }
    }


}
