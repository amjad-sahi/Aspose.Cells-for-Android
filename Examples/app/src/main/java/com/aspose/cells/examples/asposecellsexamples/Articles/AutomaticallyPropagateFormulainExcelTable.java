package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class AutomaticallyPropagateFormulainExcelTable {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Automatically Propagate Formula in Excel Table
     */
    public static void Run(Context context)
    {
        Log.w(TAG, "Running AutomaticallyPropagateFormulainExcelTable");

        try
        {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Add column headings in cell A1 and B1
            sheet.getCells().get(0, 0).putValue("Column A");
            sheet.getCells().get(0, 1).putValue("Column B");

            //Add list object, set its name and style
            ListObject listObject = sheet.getListObjects().get(sheet.getListObjects().add(0, 0, 1, sheet.getCells().getMaxColumn(), true));
            listObject.setTableStyleType(TableStyleType.TABLE_STYLE_MEDIUM_14);
            listObject.setDisplayName("Table");

            //Set the formula of second column so that it could automatically propagate to new rows while entering data
            listObject.getListColumns().get(1).setFormula("=[Column A] + 1");

            //Save the result in XLSX format
            book.save(SD_PATH + "outputAutomaticallyPropagateFormulainExcelTable.xlsx");

            Log.w(TAG, "outputAutomaticallyPropagateFormulainExcelTable.xlsx created successfully");
        }
        catch (Exception ex)
        {
            Log.e(TAG, "Some exception occurred in AutomaticallyPropagateFormulainExcelTable");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
