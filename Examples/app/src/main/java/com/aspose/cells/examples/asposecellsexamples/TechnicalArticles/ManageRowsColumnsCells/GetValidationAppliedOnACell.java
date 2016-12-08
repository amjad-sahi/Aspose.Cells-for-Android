package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Validation;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GetValidationAppliedOnACell {

    private static final String TAG = GetValidationAppliedOnACell.class.getName();

    public void getValidationAppliedOnACell() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate the workbook from sample Excel file
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Access its first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Cell C1 has the Decimal Validation applied on it.
            //It can take only the values Between 10 and 20
            Cell cell = worksheet.getCells().get("C1");

            //Access the validation applied on this cell
            Validation validation = cell.getValidation();

            //Read various properties of the validation
            Log.i(TAG, "Reading Properties of Validation");
            Log.i(TAG, "--------------------------------");
            Log.i(TAG, "Type: " + validation.getType());
            Log.i(TAG, "Operator: " + validation.getOperator());
            Log.i(TAG, "Formula1: " + validation.getFormula1());
            Log.i(TAG, "Formula2: " + validation.getFormula2());
            Log.i(TAG, "Ignore blank: " + validation.getIgnoreBlank());
        } catch (Exception e) {
            Log.e(TAG, "Get Validation Applied on a Cell", e);
        }
    }
}
