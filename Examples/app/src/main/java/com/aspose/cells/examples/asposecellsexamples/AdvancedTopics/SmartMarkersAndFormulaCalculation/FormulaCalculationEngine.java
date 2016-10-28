package com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.SmartMarkersAndFormulaCalculation;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FormulaCalculationEngine {

    private static final String TAG = FormulaCalculationEngine.class.getName();

    public void addFormulasAndCalculateResults() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook book = new Workbook();

            //Obtaining the reference of the newly added worksheet
            int sheetIndex = book.getWorksheets().add();
            Worksheet worksheet = book.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();
            Cell cell = null;

            //Adding a value to "A1" cell
            cell = cells.get("A1");
            cell.setValue(1);

            //Adding a value to "A2" cell
            cell = cells.get("A2");
            cell.setValue(2);

            //Adding a value to "A3" cell
            cell = cells.get("A3");
            cell.setValue(3);

            //Adding a SUM formula to "A4" cell
            cell = cells.get("A4");
            cell.setFormula("=SUM(A1:A3)");

            //Calculating the results of formulas
            book.calculateFormula();

            //Saving the Excel file
            book.save(filePath + File.separator + "Book1.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Formulas and Calculating Results", e);
        }
    }
}
