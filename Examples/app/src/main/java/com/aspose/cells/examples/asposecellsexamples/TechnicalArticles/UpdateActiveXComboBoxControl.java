package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ActiveXControl;
import com.aspose.cells.ComboBoxActiveXControl;
import com.aspose.cells.ControlType;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;

import java.io.File;

public class UpdateActiveXComboBoxControl {

    private static final String TAG = UpdateActiveXComboBoxControl.class.getName();

    public void updateActiveXComboBoxControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook
            Workbook wb = new Workbook(filePath + "ComboBox_Sample.xlsx");

            //Access first shape from first worksheet
            Shape shape = wb.getWorksheets().get(0).getShapes().get(0);

            //Access ActiveX ComboBox Control and update its value
            if (shape.getActiveXControl() != null) {
                //Access Shape ActiveX Control
                ActiveXControl c = shape.getActiveXControl();

                //Check if ActiveX Control is ComboBox Control
                if (c.getType() == ControlType.COMBO_BOX) {
                    //Type cast ActiveXControl into ComboBoxActiveXControl and change its value
                    ComboBoxActiveXControl comboBoxActiveX = (ComboBoxActiveXControl) c;
                    comboBoxActiveX.setValue("This is combo box control.");
                }
            }

            //Save the workbook
            wb.save(filePath + "ComboBoxControl_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Update ActiveX ComboBox Control", e);
        }
    }
}
