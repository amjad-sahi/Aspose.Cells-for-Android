package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SetTextOfChartLegendEntryFillToNone {

    private static final String TAG = SetTextOfChartLegendEntryFillToNone.class.getName();

    public void setTextOfChartLegendEntryFillToNone() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load existing spreadsheet in an instance of Workbook
            Workbook book = new Workbook(filePath + "sample-chart-legend.xlsx");

            //Access the first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Access the first chart from the sheet
            Chart chart = sheet.getCharts().get(0);

            //Set text of second legend entry fill to none
            chart.getLegend().getLegendEntries().get(1).setTextNoFill(true);

            //Save the result in xlsx format
            book.save(filePath +"SetTextOfChartLegend_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Set Text of Chart Legend Entry Fill to None", e);
        }
    }
}
