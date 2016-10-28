package com.aspose.cells.examples.asposecellsexamples.Formulas;

import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class CalculateFormulasDirectly {

    private static final String TAG = CalculateFormulasDirectly.class.getName();

    public void calculateFormulasDirectly() {
        try {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Put 20 in cell A1
            Cell cellA1 = worksheet.getCells().get("A1");
            cellA1.putValue(20);

            //Put 30 in cell A2
            Cell cellA2 = worksheet.getCells().get("A2");
            cellA2.putValue(30);

            //Calculate the Sum of A1 and A2
            Object results = worksheet.calculateFormula("=Sum(A1:A2)");

            //Print the output
            Log.v(TAG, "Value of A1: " + cellA1.getStringValue());
            Log.v(TAG, "Value of A2: " + cellA2.getStringValue());
            Log.v(TAG, "Result of Sum(A1:A2): " + results.toString());
        } catch (Exception e) {
            Log.e(TAG, "Calculate Formulas Directly", e);
        }
    }
}
