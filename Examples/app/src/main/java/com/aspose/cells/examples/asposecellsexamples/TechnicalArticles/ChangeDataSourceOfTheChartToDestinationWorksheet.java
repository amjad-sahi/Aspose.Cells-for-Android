package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CopyOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ChangeDataSourceOfTheChartToDestinationWorksheet {

    private static final String TAG = ChangeDataSourceOfTheChartToDestinationWorksheet.class.getName();

    /**
     * The following sample code explains the usage of CopyOptions.ReferToDestinationSheet property
     * while copying rows or range containing chart to new worksheet.
     */
    public void changeDataSourceOfTheChartToDestinationWorksheetWhileCopyingRowsOrRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load sample excel file
            Workbook wb = new Workbook(filePath + "ChangeDataSource.xlsx");

            //Access the first sheet which contains chart
            Worksheet source = wb.getWorksheets().get(0);

            //Add another sheet named DestSheet
            Worksheet destination = wb.getWorksheets().add("DestSheet");

            //Set CopyOptions.ReferToDestinationSheet to true
            CopyOptions options = new CopyOptions();
            options.setReferToDestinationSheet(true);

            //Copy all the rows of source worksheet to destination worksheet which includes chart as well The chart data source will now refer to DestSheet
            destination.getCells().copyRows(source.getCells(), 0, 0, source.getCells().getMaxDisplayRange().getRowCount(), options);

            //Save workbook in xlsx format
            wb.save(filePath + "ChangeDataSource_Out.xlsx", SaveFormat.XLSX);

        } catch (Exception e) {
            Log.e(TAG, "Change Data Source of the Chart to Destination Worksheet while Copying Rows or Range", e);
        }
    }
}
