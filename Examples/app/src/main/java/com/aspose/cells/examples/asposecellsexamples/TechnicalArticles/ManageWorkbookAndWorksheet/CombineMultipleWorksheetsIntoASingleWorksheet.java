package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CombineMultipleWorksheetsIntoASingleWorksheet {
    private static final String TAG = CombineMultipleWorksheetsIntoASingleWorksheet.class.getName();

    public void combineMultipleWorksheetsIntoASingleWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "source.xlsx");

            Workbook destWorkbook = new Workbook();

            Worksheet destSheet = destWorkbook.getWorksheets().get(0);

            int TotalRowCount = 0;

            for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
                Worksheet sourceSheet = workbook.getWorksheets().get(i);

                Range sourceRange = sourceSheet.getCells().getMaxDisplayRange();

                Range destRange = destSheet.getCells().createRange(sourceRange.getFirstRow() + TotalRowCount,
                        sourceRange.getFirstColumn(),
                        sourceRange.getRowCount(),
                        sourceRange.getColumnCount());

                destRange.copy(sourceRange);

                TotalRowCount = sourceRange.getRowCount() + TotalRowCount;
            }

            destWorkbook.save(filePath + "MultipleWorksheetsIntoASingle_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Combine Multiple Worksheets into a Single Worksheet", e);
        }
    }
}