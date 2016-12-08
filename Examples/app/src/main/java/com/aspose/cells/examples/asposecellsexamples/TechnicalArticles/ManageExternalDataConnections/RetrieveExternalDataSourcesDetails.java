package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ConnectionParameter;
import com.aspose.cells.ConnectionParameterCollection;
import com.aspose.cells.DBConnection;
import com.aspose.cells.ExternalConnection;
import com.aspose.cells.ExternalConnectionCollection;
import com.aspose.cells.Workbook;

import java.io.File;

public class RetrieveExternalDataSourcesDetails {

    private static final String TAG = RetrieveExternalDataSourcesDetails.class.getName();

    public void retrieveExternalDataSourceDetails() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Open the template Excel file
            Workbook workbook = new Workbook(filePath + "connection.xlsx");

            //Get the external data connections
            ExternalConnectionCollection connections = workbook.getDataConnections();
            //Get the count of the collection connection
            int connectionCount = connections.getCount();

            //Create an external connection object
            ExternalConnection connection = null;

            //Loop through all the connections in the file
            for (int i = 0; i < connectionCount; i++) {
                connection = connections.get(i);
                if (connection instanceof DBConnection) {
                    //Instantiate the DB Connection
                    DBConnection dbConn = (DBConnection) connection;

                    //Print the complete details of the object
                    Log.i(TAG, "Command: " + dbConn.getCommand());
                    Log.i(TAG, "Command Type: " + dbConn.getCommandType());
                    Log.i(TAG, "Description: " + dbConn.getConnectionDescription());
                    Log.i(TAG, "Id: " + dbConn.getConnectionId());
                    Log.i(TAG, "Info: " + dbConn.getConnectionInfo());
                    Log.i(TAG, "Credentials: " + dbConn.getCredentials());
                    Log.i(TAG, "Name: " + dbConn.getName());
                    Log.i(TAG, "OdcFile: " + dbConn.getOdcFile());
                    Log.i(TAG, "Source file: " + dbConn.getSourceFile());
                    Log.i(TAG, "Type: " + dbConn.getType());

                    //Get the parameters collection (if the connection object has)
                    ConnectionParameterCollection parameterCollection = dbConn.getParameters();
                    //Loop through all the parameters and obtain the details
                    int paramCount = parameterCollection.getCount();
                    for (int j = 0; j < paramCount; j++) {
                        ConnectionParameter param = parameterCollection.get(j);
                        Log.i(TAG, "Cell reference: " + param.getCellReference());
                        Log.i(TAG, "Parameter name: " + param.getName());
                        Log.i(TAG, "Prompt: " + param.getPrompt());
                        Log.i(TAG, "SQL Type: " + param.getSqlType());
                        Log.i(TAG, "Param Type: " + param.getType());
                        Log.i(TAG, "Param Value: " + param.getValue());
                    }
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Retrieving External Data Source Details", e);
        }
    }
}
