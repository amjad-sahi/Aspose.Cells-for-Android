package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ActiveXControl;
import com.aspose.cells.ControlType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AddActiveX {

    private static final String TAG = AddActiveX.class.getName();

    public void addToggleButtonActiveXControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook book = new Workbook();

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Add ActiveX Control of type Toggle Button to the shape collection
            Shape shape = sheet.getShapes().addActiveXControl(ControlType.TOGGLE_BUTTON, 4, 0, 4, 0, 100, 30);

            //Access the ActiveX control object and set its linked cell property
            ActiveXControl control = shape.getActiveXControl();
            control.setLinkedCell("A1");

            //Save the result in xlsx format
            book.save(filePath + "AddToggleBtnActiveXControl_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Add Toggle Button ActiveX Control", e);
        }
    }
}
