package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ExternalConnection;
import com.aspose.cells.Workbook;

import java.io.File;

import static com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections.FindQueryTablesAndListObjectsRelatedToExternalDataConnections.printTables;

public class FindReferenceCellsFromExternalConnection {

    private static final String TAG = FindReferenceCellsFromExternalConnection.class.getName();

    public void findReferenceCellsFromExternalConnection() {

        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "sample.xlsm");

            //Check all the connections inside the workbook
            for (int i = 0; i < workbook.getDataConnections().getCount(); i++) {
                ExternalConnection externalConnection = workbook.getDataConnections().get(i);
                Log.i(TAG, "connection: " + externalConnection.getName());
                printTables(workbook, externalConnection);
            }
        } catch (Exception e) {
            Log.e(TAG, "Find Reference Cells From External Connection", e);
        }
    }
}
