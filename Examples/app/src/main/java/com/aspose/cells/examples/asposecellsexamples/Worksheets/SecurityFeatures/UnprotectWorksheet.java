package com.aspose.cells.examples.asposecellsexamples.Worksheets.SecurityFeatures;

import android.util.Log;

import com.aspose.cells.Protection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class UnprotectWorksheet {
    private static final String TAG = UnprotectWorksheet.class.getName();

    public void unprotectASimplyProtectedWorksheet() {
        try {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Protection protection = worksheet.getProtection();

            //The following 3 methods are only for Excel 2000 and earlier formats
            protection.setAllowEditingContent(false);
            protection.setAllowEditingObject(false);
            protection.setAllowEditingScenario(false);

            //Un protecting the worksheet
            worksheet.unprotect();
        } catch (Exception e) {
            Log.e(TAG, "Unprotect a Simply Protected Worksheet", e);
        }
    }

    public void unprotectAPasswordProtectedWorksheet() {
        try {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Protection protection = worksheet.getProtection();

            //The following 3 methods are only for Excel 2000 and earlier formats
            protection.setAllowEditingContent(false);
            protection.setAllowEditingObject(false);
            protection.setAllowEditingScenario(false);

            //Protects the first worksheet with a password
            protection.setPassword("password");

            //Un protecting the worksheet with a password
            worksheet.unprotect("password");
        } catch (Exception e) {
            Log.e(TAG, "Unprotect a Password Protected Worksheet", e);
        }
    }
}
