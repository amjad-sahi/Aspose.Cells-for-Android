package com.aspose.cells.examples.asposecellsexamples.Chart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ChangeChartPositionAndSize {

    private static final String TAG = ChangeChartPositionAndSize.class.getName();

    /**
     *  This example loads an existing workbook that contains a chart on its first worksheet.
     *  It then re-sizes and re-positions the chart, and saves the workbook.
     */
    public void changeChartPositionAndSize() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "source.xlsx");

            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Load the chart from source worksheet
            Chart chart = worksheet.getCharts().get(0);

            //Resize the chart
            chart.getChartObject().setWidth(400);
            chart.getChartObject().setHeight(300);

            //Reposition the chart
            chart.getChartObject().setX(250);
            chart.getChartObject().setY(150);

            //Output the file
            workbook.save(filePath + File.separator + "ChangeChartPositionAndSize_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Change the Chart's Position and Size", e);
        }
    }
}
