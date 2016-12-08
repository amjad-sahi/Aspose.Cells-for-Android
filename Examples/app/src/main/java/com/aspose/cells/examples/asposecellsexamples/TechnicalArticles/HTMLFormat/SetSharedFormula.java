package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class SetSharedFormula {
    private static final String TAG = SetSharedFormula.class.getName();

    public void setSharedFormula() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a Workbook from existing file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Get the cells collection in the first worksheet
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Apply the shared formula in the range i.e.., B2:B14
            cells.get("B2").setSharedFormula("=A2*0.09", 13, 1);

            //Save the Excel file
            workbook.save(filePath + "SetSharedFormula_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Set Shared Formula", e);
        }
    }

}
