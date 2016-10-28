package com.aspose.cells.examples.asposecellsexamples.Formulas;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CalculateFormulas {
    private static final String TAG = CalculateFormulas.class.getName();

    public void addFormulasAndCalculateResults() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Excel object
            int sheetIndex = workbook.getWorksheets().add();

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Adding a value to "A1" cell
            worksheet.getCells().get("A1").putValue(1);

            //Adding a value to "A2" cell
            worksheet.getCells().get("A2").putValue(2);

            //Adding a value to "A3" cell
            worksheet.getCells().get("A3").putValue(3);

            //Adding a SUM formula to "A4" cell
            worksheet.getCells().get("A4").setFormula("=SUM(A1:A3)");

            //Calculating the results of formulas
            workbook.calculateFormula();

            //Get the calculated value of the cell
            String value = worksheet.getCells().get("A4").getStringValue();

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddFormulas_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Formulas and Calculate Results", e);
        }
    }
}
