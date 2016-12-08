package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class AdjustRowsAndColumns {

    private static final String TAG = AdjustRowsAndColumns.class.getName();

    public void adjustRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            int fileFormatType = FileFormatType.EXCEL_97_TO_2003;
            Workbook workbook = new Workbook(fileFormatType);
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            Cells cells = worksheet.getCells();

            //Set the height of all row in the worksheet
            cells.setStandardHeight(20);
            //Set the width of all columns in the worksheet
            cells.setStandardWidth(20);

            //Set the width of the first column
            cells.setColumnWidth(0, 12);
            //Set the width of the second column
            cells.setColumnWidth(1, 40);
            //Setting the height of row
            cells.setRowHeight(1, 8);
            workbook.save(filePath + "Cells_AdjustingRowsAndColumns_Out.xls", SaveFormat.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Adjusting Rows And Columns", e);
        }
    }
}
