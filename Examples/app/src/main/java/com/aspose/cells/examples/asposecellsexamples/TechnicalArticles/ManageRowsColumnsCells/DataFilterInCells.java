package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class DataFilterInCells {

    private static final String TAG = DataFilterInCells.class.getName();

    public void dataFilterInCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook();

            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Cells cells = worksheet.getCells();
            Cell cell;

            //Put a value into a cell
            cell = cells.get("A1");
            cell.setValue("Fruit");
            cell = cells.get("B1");
            cell.setValue("Total");
            cell = cells.get("A2");
            cell.setValue("Apple");
            cell = cells.get("B2");
            cell.setValue(1000);
            cell = cells.get("A3");
            cell.setValue("Orange");
            cell = cells.get("B3");
            cell.setValue(2500);
            cell = cells.get("A4");
            cell.setValue("Bananas");
            cell = cells.get("B4");
            cell.setValue(2500);
            cell = cells.get("A5");
            cell.setValue("Pear");
            cell = cells.get("B5");
            cell.setValue(1000);
            cell = cells.get("A6");
            cell.setValue("Grape");
            cell = cells.get("B6");
            cell.setValue(2000);

            cell = cells.get("D1");
            cell.setValue("Count:");
            cell = cells.get("E1");
            cell.setFormula("=SUBTOTAL(2, B1:B6)");

            worksheet.getAutoFilter().setRange("A1:B6");

            workbook.save(filePath + "DataFilter_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Data Filter in Cells", e);
        }
    }
}
