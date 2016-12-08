package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.QueryTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ReadAndWriteQueryTableOfWorksheet {

    private static final String TAG = ReadAndWriteQueryTableOfWorksheet.class.getName();

    public void readAndWriteQueryTableOfWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook from source excel file
            Workbook workbook = new Workbook(filePath + "QueryTable.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first Query Table
            QueryTable qt = worksheet.getQueryTables().get(0);

            //Print Query Table Data
            System.out.println("Adjust Column Width: " + qt.getAdjustColumnWidth());
            System.out.println("Preserve Formatting: " + qt.getPreserveFormatting());

            //Now set Preserve Formatting to true
            qt.setPreserveFormatting(true);

            //Save the workbook
            workbook.save(filePath + "ReadAndWriteQueryTable_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Reading and Writing Query Table of Worksheet", e);
        }
    }
}
