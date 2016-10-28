package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import java.io.File;

public class AddDataToCells {
    private static final String TAG = AddDataToCells.class.getName();

    public void addDataToCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook();
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding a string value to the cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello World");

            //Adding a double value to the cell
            cell = cells.get("A2");
            cell.setValue(20.5);

            //Adding an integer  value to the cell
            cell = cells.get("A3");
            cell.setValue(15);

            //Adding a boolean value to the cell
            cell = cells.get("A4");
            cell.setValue(true);

            //Adding a date/time value to the cell
            cell = cells.get("A5");
            DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
            Date date = new Date();
            cell.setValue(dateFormat.format(date));

            //Setting the display format of the date
            Style style = cell.getStyle();
            style.setNumber(15);
            cell.setStyle(style);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddDataToCells_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Data to Cells", e);
        }
    }


}
