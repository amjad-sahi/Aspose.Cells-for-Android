package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageVBAModules;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class CheckIfVBACodeIsSigned {
    private static final String TAG = CheckIfVBACodeIsSigned.class.getName();

    public void checkIfVBACodeIsSigned() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "sampleVBAProjectSigned.xlsm");

            Log.i(TAG, "Is VBA Code Project Signed: " + workbook.getVbaProject().isSigned());
        } catch (Exception e) {
            Log.e(TAG, "Check if VBA Code is Signed", e);
        }
    }
}
