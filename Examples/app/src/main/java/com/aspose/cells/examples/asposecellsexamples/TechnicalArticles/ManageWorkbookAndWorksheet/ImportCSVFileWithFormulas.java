package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ImportCSVFileWithFormulas {

    private static final String TAG = ImportCSVFileWithFormulas.class.getName();

    public void importCSVFileWithFormulas() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            String csvFile = filePath + "sample.csv";

            TxtLoadOptions opts = new TxtLoadOptions();
            opts.setSeparator(',');
            opts.setHasFormula(true);

            //Load your CSV file with formulas in a Workbook object
            Workbook workbook = new Workbook(csvFile, opts);

            //You can also import your CSV file like this
            //The code below is importing CSV file starting from cell D4
            Worksheet worksheet = workbook.getWorksheets().get(0);
            worksheet.getCells().importCSV(csvFile, opts, 3, 3);

            //Save your workbook in Xlsx format
            workbook.save(filePath + "ImportCSVFile_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Load or Import CSV file with Formulas", e);
        }
    }
}
