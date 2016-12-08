package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SaveEachWorksheetToADifferentPDFFile {
    private static final String TAG = SaveEachWorksheetToADifferentPDFFile.class.getName();

    public void saveEachWorksheetToADifferentPDFFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Get the count of the worksheets in the workbook
            int sheetCount = workbook.getWorksheets().getCount();

            //Make all sheets invisible except first worksheet
            for (int i = 1; i < workbook.getWorksheets().getCount(); i++) {
                workbook.getWorksheets().get(i).setVisible(false);
            }

            //Save PDF for each worksheet
            for (int j = 0; j < workbook.getWorksheets().getCount(); j++) {
                Worksheet ws = workbook.getWorksheets().get(j);
                workbook.save(filePath + ws.getName() + "_Out.pdf");

                if (j < workbook.getWorksheets().getCount() - 1) {
                    workbook.getWorksheets().get(j + 1).setVisible(true);
                    workbook.getWorksheets().get(j).setVisible(false);
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Save Each Worksheet to a Different PDF File", e);
        }
    }
}
