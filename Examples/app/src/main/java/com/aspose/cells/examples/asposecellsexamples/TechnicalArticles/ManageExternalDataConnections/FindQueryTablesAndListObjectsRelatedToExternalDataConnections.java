package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.ExternalConnection;
import com.aspose.cells.ListObject;
import com.aspose.cells.Name;
import com.aspose.cells.QueryTable;
import com.aspose.cells.Range;
import com.aspose.cells.TableDataSourceType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FindQueryTablesAndListObjectsRelatedToExternalDataConnections {

    private static final String TAG = FindQueryTablesAndListObjectsRelatedToExternalDataConnections.class.getName();

    public void findQueryTablesAndListObjects() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load workbook object
            Workbook workbook = new Workbook(filePath + "sample.xlsm");

            //Check all the connections inside the workbook
            for (int i = 0; i < workbook.getDataConnections().getCount(); i++) {
                ExternalConnection externalConnection = workbook.getDataConnections().get(i);
                Log.i(TAG, "connection: " + externalConnection.getName());
                printTables(workbook, externalConnection);
            }
        } catch (Exception e) {
            Log.e(TAG, "Find Query Tables and List Objects related to External Data Connections", e);
        }
    }

    public static void printTables(Workbook workbook, ExternalConnection ec) {
        //Iterate all the worksheets
        for (int j = 0; j < workbook.getWorksheets().getCount(); j++) {
            Worksheet worksheet = workbook.getWorksheets().get(j);

            //Check all the query tables in a worksheet
            for (int k = 0; k < worksheet.getQueryTables().getCount(); k++) {
                QueryTable qt = worksheet.getQueryTables().get(k);

                //Check if query table is related to this external connection
                if (ec.getId() == qt.getConnectionId()
                        && qt.getConnectionId() >= 0) {
                    //Print the query table name and print its "Refers To" range
                    Log.i(TAG, "querytable " + qt.getName());
                    String n = qt.getName().replace('+', '_').replace('=', '_');
                    Name name = workbook.getWorksheets().getNames().get("'" + worksheet.getName() + "'!" + n);
                    if (name != null) {
                        Range range = name.getRange();
                        if (range != null) {
                            Log.i(TAG, "Refers To: " + range.getRefersTo());
                        }
                    }
                }
            }

            //Iterate all the list objects in this worksheet
            for (int k = 0; k < worksheet.getListObjects().getCount(); k++) {
                ListObject table = worksheet.getListObjects().get(k);

                //Check the data source type if it is query table
                if (table.getDataSourceType() == TableDataSourceType.QUERY_TABLE) {
                    //Access the query table related to list object
                    QueryTable qt = table.getQueryTable();

                    //Check if query table is related to this external connection
                    if (ec.getId() == qt.getConnectionId()
                            && qt.getConnectionId() >= 0) {
                        //Print the query table name and print its refersto range
                        Log.i(TAG, "querytable " + qt.getName());
                        Log.i(TAG, "Table " + table.getDisplayName());
                        Log.i(TAG, "refersto: " + worksheet.getName() + "!" + CellsHelper.cellIndexToName(table.getStartRow(), table.getStartColumn()) + ":" + CellsHelper.cellIndexToName(table.getEndRow(), table.getEndColumn()));
                    }
                }
            }
        }
    }
}