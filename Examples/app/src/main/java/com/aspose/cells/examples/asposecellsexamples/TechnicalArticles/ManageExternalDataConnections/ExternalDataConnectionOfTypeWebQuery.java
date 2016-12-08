package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ExternalConnection;
import com.aspose.cells.WebQueryConnection;
import com.aspose.cells.Workbook;

import java.io.File;

public class ExternalDataConnectionOfTypeWebQuery {
    private static final String TAG = ExternalDataConnectionOfTypeWebQuery.class.getName();

    public void workWithExternalDataConnectionOfTypeWebQuery() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "WebQuerySample.xlsx");
            ExternalConnection connection = workbook.getDataConnections().get(0);

            if (connection instanceof WebQueryConnection) {
                WebQueryConnection webQuery = (WebQueryConnection) connection;
                Log.i(TAG, "Web Query URL: " + webQuery.getUrl());
            }
        } catch (Exception e) {
            Log.e(TAG, "Work with External Data Connection of type WebQuery", e);
        }
    }
}
