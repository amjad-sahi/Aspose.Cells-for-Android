package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.DataLabels;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ResizeChartsDataLabelShapeToFitText {

    private static final String TAG = ResizeChartsDataLabelShapeToFitText.class.getName();

    public void resizeChartsDataLabelShapeToFitText() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook containing the Chart
            Workbook book = new Workbook(filePath + "ChangeDataSource.xlsx");

            //Access the Worksheet that contains the Chart
            Worksheet sheet = book.getWorksheets().get(0);

            //Access ChartCollection from Worksheet
            ChartCollection charts = sheet.getCharts();

            //Loop over each chart in collection
            for (int chartIndex = 0; chartIndex < charts.getCount(); chartIndex++) {
                //Access indexed chart from the collection
                Chart chart = charts.get(chartIndex);

                for (int seriesIndex = 0; seriesIndex < chart.getNSeries().getCount(); seriesIndex++) {
                    //Access the DataLabels of indexed NSeries
                    DataLabels labels = chart.getNSeries().get(seriesIndex).getDataLabels();

                    //Set ResizeShapeToFitText property to true
                    labels.setResizeShapeToFitText(true);
                }

                //Calculate Chart
                chart.calculate();
            }

            //Save the result
            book.save(filePath + "ResizeShapeToFitText_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Resize Chart's Data Label Shape To Fit Text", e);
        }
    }
}
