package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FontSetting;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;

import java.io.File;
import java.util.ArrayList;

public class ChangeCharacterSpacingOfExcelTextBoxOrShape {

    private static final String TAG = ChangeCharacterSpacingOfExcelTextBoxOrShape.class.getName();

    public void changeCharacterSpacingOfExcelTextBoxOrShape() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load existing spreadsheet in an instance of Workbook
            Workbook book = new Workbook(filePath +  "sample-character-spacing.xlsx");

            //Access text box which is also a shape object
            Shape shape = book.getWorksheets().get(0).getShapes().get(0);

            //Access the first font setting object
            ArrayList<FontSetting> list = shape.getCharacters();
            FontSetting setting = list.get(0);

            //Set the character spacing to 4
            setting.getShapeFont().setSpacing(4);

            //Save the result in xlsx format
            book.save(filePath + "ChangeCharacterSpacing_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Change Character Spacing of Excel TextBox or Shape", e);
        }
    }
}