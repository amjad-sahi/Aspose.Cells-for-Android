package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ImportXMLMapInsideAWorkbook {

    private static final String TAG = ImportXMLMapInsideAWorkbook.class.getName();

    public void importXMLMapInsideAWorkbook () {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook
            Workbook workbook = new Workbook();

            //URL that contains your XML data for mapping
            String XML = "http://www.aspose.com/docs/download/attachments/434475650/sampleXML.txt";

            //Import your XML Map data starting from cell A1
            workbook.importXml(XML, "Sheet1", 0, 0);

            //Save workbook
            workbook.save(filePath + "ImportXMLMap_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Import XML Map inside a Workbook", e);
        }
    }

}
