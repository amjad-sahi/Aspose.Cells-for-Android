package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.ByteArrayOutputStream;
import java.io.File;

public class ConvertChartToPDF {

    private static final String TAG = ConvertChartToPDF.class.getName();

    public void convertChartToPDF() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "sample.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first chart inside the worksheet
            Chart chart = worksheet.getCharts().get(0);

            //Save the chart into pdf format
            chart.toPdf(filePath + File.separator + "Chart_Out.pdf");
        } catch (Exception e) {
            Log.e(TAG, "Convert Chart to PDF", e);
        }
    }

    public void saveChartPDFInByteArrayOutputStream() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "sample.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first chart inside the worksheet
            Chart chart = worksheet.getCharts().get(0);

            //Save the chart to PDF as Stream
            ByteArrayOutputStream outStream = new ByteArrayOutputStream();
            chart.toPdf(outStream);
        } catch (Exception e) {
            Log.e(TAG, "Save Chart PDF In ByteArrayOutputStream", e);
        }
    }


}
