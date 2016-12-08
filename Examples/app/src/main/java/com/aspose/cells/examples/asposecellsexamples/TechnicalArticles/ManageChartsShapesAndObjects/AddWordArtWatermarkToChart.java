package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.MsoFillFormat;
import com.aspose.cells.MsoPresetTextEffect;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;

import java.io.File;

public class AddWordArtWatermarkToChart {

    private static final String TAG = AddWordArtWatermarkToChart.class.getName();

    public void addWordArtWatermarkToChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new workbook.
            //Open the existing Excel file.
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Get the chart in the first worksheet.
            Chart chart = workbook.getWorksheets().get(0).getCharts().get(0);

            //Add a WordArt watermark (shape) to the chart's plot area.
            Shape wordart = chart.getShapes().addTextEffectInChart(MsoPresetTextEffect.TEXT_EFFECT_2,
                                                        "CONFIDENTIAL", "Arial Black", 66, false, false,
                                                        1200, 500, 2000, 3000);

            //Get the shape's fill format.
            MsoFillFormat wordArtFormat = wordart.getFillFormat();

            //Set the transparency.
            wordArtFormat.setTransparency(0.9);

            //Get the line format and make it invisible.
            com.aspose.cells.MsoLineFormat lineFormat = wordart.getLineFormat();
            lineFormat.setVisible(false);

            //Save the Excel file.
            workbook.save(filePath + "AddWordArtWatermarkToChart_Out.xls", SaveFormat.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Add WordArt Watermark to Chart", e);
        }
    }
}