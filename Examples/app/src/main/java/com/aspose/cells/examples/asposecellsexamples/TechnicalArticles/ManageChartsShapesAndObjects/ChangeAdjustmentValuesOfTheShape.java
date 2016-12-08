package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ChangeAdjustmentValuesOfTheShape {

    private static final String TAG = ChangeAdjustmentValuesOfTheShape.class.getName();

    public void changeAdjustmentValuesOfTheShape() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first three shapes of the worksheet
            Shape shape1 = worksheet.getShapes().get(0);
            Shape shape2 = worksheet.getShapes().get(1);
            Shape shape3 = worksheet.getShapes().get(2);

            //Change the adjustment values of the shapes
            shape1.getGeometry().getShapeAdjustValues().get(0).setValue(0.5d);
            shape2.getGeometry().getShapeAdjustValues().get(0).setValue(0.8d);
            shape3.getGeometry().getShapeAdjustValues().get(0).setValue(0.5d);

            //Save the workbook
            workbook.save(filePath + "ShapreAdjustValues_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Change Adjustment Values of the Shape", e);
        }
    }

}
