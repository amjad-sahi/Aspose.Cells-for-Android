package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.ErrorCheckOption;
import com.aspose.cells.ErrorCheckOptionCollection;
import com.aspose.cells.ErrorCheckType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class UseErrorCheckingOptions {

    private static final String TAG = UseErrorCheckingOptions.class.getName();

    public void disableTheNumbersStoredAsTextErrorCheckingOption() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook and opening a template spreadsheet
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the first worksheet
            Worksheet sheet = workbook.getWorksheets().get(0);
            //Instantiate the error checking options
            ErrorCheckOptionCollection opts = sheet.getErrorCheckOptions();

            int index = opts.add();
            ErrorCheckOption opt = opts.get(index);
            //Disable the numbers stored as text option
            opt.setErrorCheck(ErrorCheckType.TEXT_NUMBER, false);
            //Set the range
            opt.addRange(CellArea.createCellArea(0, 0, 65535, 255));

            //Save the Excel file
            workbook.save(filePath + "ErrorCheckingOptions_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Use Error Checking Options", e);
        }
    }
}
