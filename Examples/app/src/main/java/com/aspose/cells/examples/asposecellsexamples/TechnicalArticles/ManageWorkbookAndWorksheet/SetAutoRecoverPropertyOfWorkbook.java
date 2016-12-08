package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class SetAutoRecoverPropertyOfWorkbook {

    private static final String TAG = SetAutoRecoverPropertyOfWorkbook.class.getName();

    public void setAutoRecoverPropertyOfWorkbook() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook();

            //Read AutoRecover property
            Log.i(TAG, "AutoRecover: " + workbook.getSettings().getAutoRecover());

            //Set AutoRecover property to false
            workbook.getSettings().setAutoRecover(false);

            //Save the workbook
            workbook.save(filePath + "AutoRecoveryProperty_Out.xlsx");

            //Read the saved workbook again
            workbook = new Workbook(filePath + "AutoRecoveryProperty_Out1.xlsx");

            //Read AutoRecover property
            Log.i(TAG, "AutoRecover: " + workbook.getSettings().getAutoRecover());
        } catch (Exception e) {
            Log.e(TAG, "Set AutoRecover property of Workbook", e);
        }
    }
}
