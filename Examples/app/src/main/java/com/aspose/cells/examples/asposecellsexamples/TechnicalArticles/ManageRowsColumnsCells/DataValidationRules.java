package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DataValidationRules {

    private static final String TAG = DataValidationRules.class.getName();

    public void verifyThatCellValueSatisfiesDataValidationRules() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate the workbook from sample Excel file
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Access the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access Cell C1
            //Cell C1 has the Decimal Validation applied on it.
            //It can take only the values Between 10 and 20
            Cell cell = worksheet.getCells().get("C1");

            //Enter 3 inside this cell
            //Since it is not between 10 and 20, it should fail the validation
            cell.putValue(3);

            //Check if number 3 satisfies the Data Validation rule applied on this cell
            Log.e(TAG, "Is 3 a Valid Value for this Cell: " + cell.getValidationValue());

            //Enter 15 inside this cell
            //Since it is between 10 and 20, it should succeed the validation
            cell.putValue(15);

            //Check if number 15 satisfies the Data Validation rule applied on this cell
            Log.e(TAG, "Is 15 a Valid Value for this Cell: " + cell.getValidationValue());

            //Enter 30 inside this cell
            //Since it is not between 10 and 20, it should fail the validation again
            cell.putValue(30);

            //Check if number 30 satisfies the Data Validation rule applied on this cell
            Log.e(TAG, "Is 30 a Valid Value for this Cell: " + cell.getValidationValue());
        } catch (Exception e) {
            Log.e(TAG, "Get Validation Applied on a Cell", e);
        }
    }
}
