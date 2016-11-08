package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.SmartMarkers;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartPoint;
import com.aspose.cells.Series;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FindIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart {

    private static final String TAG = FindIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart.class.getName();

    public void findIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load source excel file inside the workbook object
            //Get the SD card path
            String dirPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

            //Load source excel file containing Bar of Pie chart
            Workbook wb = new Workbook(dirPath + "PieBars.xlsx");

            //Access first worksheet
            Worksheet ws = wb.getWorksheets().get(0);

            //Access first chart which is Bar of Pie chart and calculate it
            Chart ch = ws.getCharts().get(0);
            ch.calculate();

            //Access the chart series
            Series srs = ch.getNSeries().get(0);

            //Print the data points of the chart series and check
            //its IsInSecondaryPlot property to determine if data point is inside the bar or pie
            for(int i=0; i<srs.getPoints().getCount(); i++) {
                //Access chart point
                ChartPoint cp = srs.getPoints().get(i);

                //Skip null values
                if (cp.getYValue() == null)
                    continue;

                //Print the chart point value and see if it is inside bar or pie
                //If the IsInSecondaryPlot is true, then the data point is inside bar
                //otherwise it is inside the pie
                Log.i(TAG, "Value: " + cp.getYValue());
                Log.i(TAG, "IsInSecondaryPlot: " + cp.isInSecondaryPlot());
                Log.i(TAG, "");
            }

        } catch (Exception e) {
            Log.e(TAG, "Find if Data Points are in the Second Pie or Bar on a Pie of Pie or Bar of Pie Chart", e);
        }
    }
}
