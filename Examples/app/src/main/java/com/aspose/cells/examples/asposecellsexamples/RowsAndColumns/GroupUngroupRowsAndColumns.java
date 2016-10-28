package com.aspose.cells.examples.asposecellsexamples.RowsAndColumns;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GroupUngroupRowsAndColumns {

    private static final String TAG = GroupUngroupRowsAndColumns.class.getName();

    public void groupRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Grouping first six rows (from 0 to 5) and making them hidden by passing true
            cells.groupRows(0,5,true);

            //Grouping first three columns (from 0 to 2) and making them hidden by passing true
            cells.groupColumns(0,2,true);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "GroupRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Group Rows And Columns", e);
        }
    }

    public void ungroupRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Ungrouping first six rows (from 0 to 5)
            cells.ungroupRows(0,5);

            //Ungrouping first three columns (from 0 to 2)
            cells.ungroupColumns(0,2);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "UngroupRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "UnGroup Rows And Columns", e);
        }
    }

    public void summaryRowsBelowDetail() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //SummaryRowBelow controls whether there's a summary row below the group.
            worksheet.getOutline().SummaryRowBelow = false;

            //SummaryColumnRight controls whether there's a summary row to the right of the group.
            worksheet.getOutline().SummaryColumnRight = false;
        } catch (Exception e) {
            Log.e(TAG, "Summary Rows Below Detail", e);
        }
    }
}