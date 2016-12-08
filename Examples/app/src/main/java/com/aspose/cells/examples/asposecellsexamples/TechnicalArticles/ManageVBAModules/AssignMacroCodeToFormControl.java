package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageVBAModules;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Color;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.PlacementType;
import com.aspose.cells.VbaModule;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AssignMacroCodeToFormControl {

    private static final String TAG = AssignMacroCodeToFormControl.class.getName();

    public void assignMacroCodeToFormControl() {

        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.getWorksheets().get(0);

            int moduleIdx = workbook.getVbaProject().getModules().add(sheet);
            VbaModule module = workbook.getVbaProject().getModules().get(moduleIdx);
            module.setCodes("Sub ShowMessage()" + "\r\n" + "    MsgBox \"Welcome to Aspose!\"" + "\r\n" + "End Sub");

            com.aspose.cells.Button button = (com.aspose.cells.Button) sheet.getShapes().addShape(MsoDrawingType.BUTTON, 2, 0, 2, 0, 28, 80);
            button.setPlacement(PlacementType.FREE_FLOATING);
            button.getFont().setName("Tahoma");
            button.getFont().setBold(true);
            button.getFont().setColor(Color.getBlue());
            button.setText("Aspose");
            button.setMacroName(sheet.getName() + ".ShowMessage");

            workbook.save(filePath + "AssignMacroCodeToFormControl_Out.xlsm");
        } catch (Exception e) {
            Log.e(TAG, "Assign Macro Code to Form Control", e);
        }
    }
}
