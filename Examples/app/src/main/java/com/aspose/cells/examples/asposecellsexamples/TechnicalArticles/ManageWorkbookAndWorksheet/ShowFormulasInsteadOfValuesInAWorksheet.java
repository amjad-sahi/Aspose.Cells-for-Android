package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ShowFormulasInsteadOfValuesInAWorksheet {

    private static final String TAG = ShowFormulasInsteadOfValuesInAWorksheet.class.getName();

    public void showFormulasInsteadOfValuesInAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load the source workbook
            Workbook workbook = new Workbook(filePath + "ShowFormulas.xlsx");

            //Access the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Show formulas of the worksheet
            worksheet.setShowFormulas(true);

            //Save the workbook
            workbook.save(filePath + "ShowFormulas_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Show Formulas instead of Values in a Worksheet", e);
        }
    }
}
