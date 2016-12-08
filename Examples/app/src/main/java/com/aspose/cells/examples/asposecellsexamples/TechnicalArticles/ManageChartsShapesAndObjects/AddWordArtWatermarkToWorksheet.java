package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.FillType;
import com.aspose.cells.MsoPresetTextEffect;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AddWordArtWatermarkToWorksheet {

    private static final String TAG = AddWordArtWatermarkToWorksheet.class.getName();

    public void addWordArtWatermarkToWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new Workbook
            Workbook workbook = new Workbook();

            //Get the first default sheet
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Add Watermark
            Shape wordart = sheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1,
                    "CONFIDENTIAL", "Arial Black", 50, false, true, 18, 8, 1, 1, 130, 800);

            //Get the fill format of the word art
            FillFormat wordArtFormat = wordart.getFill();

            //Set the color
            wordArtFormat.setFillType(FillType.SOLID);
            wordArtFormat.getSolidFill().setColor(Color.getRed());

            //Set the transparency
            wordArtFormat.setTransparency(0.9);

            //Make the line invisible
            wordart.setHasLine(false);

            //Save the file
            workbook.save(filePath + "WatermarkToWorksheet_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add WordArt Watermark to Worksheet", e);
        }
    }
}
