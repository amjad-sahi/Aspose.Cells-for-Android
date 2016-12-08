package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class Implement1904DateSystem {

    private static final String TAG = Implement1904DateSystem.class.getName();

    public void implement1904DateSystem() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Initialize a new Workbook
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Implement 1904 date system
            workbook.getSettings().setDate1904(true);

            //Save the Excel file
            workbook.save(filePath + "Implement1904DateSystem_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Implement 1904 Date System", e);
        }
    }

}
