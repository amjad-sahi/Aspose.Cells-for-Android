package com.aspose.cells.examples.asposecellsexamples.Formulas;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CalculateFormulasOnceOnly {

    private static final String TAG = CalculateFormulasOnceOnly.class.getName();

    /**
     * To improve Aspose.Cells' formula calculation performance and avoid creating a formula calculating chain,
     * set Workbook.getSettings().setCreateCalcChain() to false. It is true by default.
     */
    public void calculateFormulasOnceOnly() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Load the template workbook
            Workbook workbook = new Workbook(filePath + File.separator + "source.xlsx");

            //Print the time before formula calculation
            DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
            Date date = new Date();
            Log.v(TAG, dateFormat.format(date));

            //Set the CreateCalcChain as false
            workbook.getSettings().setCreateCalcChain(false);

            //Calculate the workbook formulas
            workbook.calculateFormula();

            //Print the time after formula calculation
            date = new Date();
            Log.v(TAG, dateFormat.format(date));
        } catch (Exception e) {
            Log.e(TAG, "Calculate Formulas Once Only", e);
        }
    }
}
