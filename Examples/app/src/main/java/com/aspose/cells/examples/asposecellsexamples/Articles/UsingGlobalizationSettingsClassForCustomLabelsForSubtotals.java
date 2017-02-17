package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.*;

public class UsingGlobalizationSettingsClassForCustomLabelsForSubtotals {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Using Globalization Settings Class For Custom Labels For Subtotals
     */
    public static void Run(Context context) {
        UsingGlobalizationSettingsClassForCustomLabelsForSubtotals pg = new UsingGlobalizationSettingsClassForCustomLabelsForSubtotals();
        pg.RunUsingGlobalizationSettingsClassForCustomLabelsForSubtotals(context);
    }

    public class CustomSettings extends GlobalizationSettings
    {
        public String getTotalName(int functionType)
        {
            switch (functionType)
            {
                case ConsolidationFunction.AVERAGE:
                    return "AVG";
                default:
                    return super.getTotalName(functionType);
            }
        }

        public String getGrandTotalName(int functionType)
        {
            switch (functionType)
            {
                case ConsolidationFunction.AVERAGE:
                    return "GRAND AVG";
                default:
                    return super.getGrandTotalName(functionType);
            }

        }
    }

    public void RunUsingGlobalizationSettingsClassForCustomLabelsForSubtotals(Context context) {

        Log.w(TAG, "Running UsingGlobalizationSettingsClassForCustomLabelsForSubtotals");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleCustomLabelsForSubtotals.xlsx");

            //Load a sample spreadsheet containing a shape
            Workbook book = new Workbook(in);

            //Assigns the GlobalizationSettings property of the WorkbookSettings class
            //to the class created in first step
            book.getSettings().setGlobalizationSettings(new CustomSettings());

            //Accesses the 1st worksheet from the collection which contains data
            //Data resides in the cell range A2:B9
            Worksheet sheet = book.getWorksheets().get(0);

            //Adds SubTotal of type Average to the worksheet
            sheet.getCells().subtotal(CellArea.createCellArea("A2", "B9"), 0, ConsolidationFunction.AVERAGE, new int[] { 0,1 });

            //Calculates Formulas
            book.calculateFormula();

            //Auto fits all columns
            sheet.autoFitColumns();

            //Save the result in XLSX format
            book.save(SD_PATH + "outputCustomLabelsForSubtotals.xlsx");

            Log.w(TAG, "outputCustomLabelsForSubtotals.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in UsingGlobalizationSettingsClassForCustomLabelsForSubtotals");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }
    }

}
