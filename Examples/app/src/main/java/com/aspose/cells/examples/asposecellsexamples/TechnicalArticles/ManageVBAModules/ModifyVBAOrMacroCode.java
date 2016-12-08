package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageVBAModules;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.VbaModule;
import com.aspose.cells.VbaModuleCollection;
import com.aspose.cells.Workbook;

import java.io.File;

public class ModifyVBAOrMacroCode {

    private static final String TAG = ModifyVBAOrMacroCode.class.getName();

    public void modifyVBAOrMacroCode() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source Excel file
            Workbook workbook = new Workbook(filePath + "sample.xlsm");

            //Change the VBA Module Code
            VbaModuleCollection modules = workbook.getVbaProject().getModules();

            for (int i = 0; i < modules.getCount(); i++) {
                VbaModule module = modules.get(i);
                String code = module.getCodes();

                //Replace the original message with the modified message
                if (code.contains("This is test message.")) {
                    code = code.replace("This is test message.", "This is Aspose.Cells message.");
                    module.setCodes(code);
                }
            }

            //Save the output Excel file
            workbook.save(filePath + "ModifyVBAOrMacroCode_Out.xlsm");
        } catch (Exception e) {
            Log.e(TAG, "Modifying VBA or Macro Code", e);
        }
    }
}
