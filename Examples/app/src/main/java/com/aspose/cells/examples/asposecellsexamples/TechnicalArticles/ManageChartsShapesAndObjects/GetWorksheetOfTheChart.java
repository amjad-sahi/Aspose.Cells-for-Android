package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GetWorksheetOfTheChart {

    private static final String TAG = GetWorksheetOfTheChart.class.getName();

    public void getWorksheetOfTheChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook from sample Excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet of the workbook
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Print worksheet name
            Log.i(TAG, "Sheet Name: " + worksheet.getName());

            //Access the first chart inside this worksheet
            Chart chart = worksheet.getCharts().get(0);

            //Access the chart's sheet and display its name again
            Log.i(TAG, "Chart's Sheet Name: " + chart.getWorksheet().getName());
        } catch (Exception e) {
            Log.e(TAG, "Get Validation Applied on a Cell", e);
        }
    }
}
