package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class VerifyPasswordUsedToProtectTheWorksheet {

    private static final String TAG = VerifyPasswordUsedToProtectTheWorksheet.class.getName();

    public void verifyPassword() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load a spreadsheet
            Workbook book = new Workbook(filePath + "sample.xlsx");

            //Access the protected Worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Check if Worksheet is password protected
            if (sheet.getProtection().isProtectedWithPassword()) {
                //Verify the password used to protect the Worksheet
                if (sheet.getProtection().verifyPassword("password")) {
                    System.out.println("Specified password has matched");
                }
                else {
                    System.out.println("Specified password has not matched");
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Verify Password Used to Protect the Worksheet", e);
        }
    }

}
