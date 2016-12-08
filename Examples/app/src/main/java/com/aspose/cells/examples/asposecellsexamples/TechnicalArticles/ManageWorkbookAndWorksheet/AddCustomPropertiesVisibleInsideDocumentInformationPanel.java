package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;

import java.io.File;

public class AddCustomPropertiesVisibleInsideDocumentInformationPanel {

    private static final String TAG = AddCustomPropertiesVisibleInsideDocumentInformationPanel.class.getName();

    public void addCustomPropertiesVisibleInsideDocumentInformationPanel() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook(FileFormatType.XLSX);

            //Add simple property without any type
            workbook.getContentTypeProperties().add("MK31", "Simple Data");

            //Add date time property with type
            workbook.getContentTypeProperties().add("MK32", "04-Mar-2015", "DateTime");

            //Save the workbook
            workbook.save(filePath + "AddCustomProperties_Out.xlsx");

        } catch (Exception e) {
            Log.e(TAG, "Add Custom Properties visible inside Document Information Panel", e);
        }
    }

}
