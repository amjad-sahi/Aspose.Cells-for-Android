package com.example;

import java.io.InputStream;

import com.aspose.cells.*;

import android.os.Bundle;
import android.app.Activity;
import android.content.res.AssetManager;
import android.view.Menu;
import android.widget.TextView;


public class FindReferenceCellsFromExternalConnectionActivity extends Activity {

	public void findReferenceCellsFromExternalConnection() throws Exception
	{	
		//Load workbook object found in assets	
		AssetManager am = getApplicationContext().getAssets();
		InputStream is = am.open("sample.xlsm");
		Workbook workbook = new Workbook(is);
		
		//Check all the connections inside the workbook
		for (int i = 0; i < workbook.getDataConnections().getCount(); i++)
		{
			ExternalConnection externalConnection = workbook.getDataConnections().get(i);
			System.out.println("connection: " + externalConnection.getName());
			PrintTables(workbook, externalConnection);
			System.out.println();
		}
		
	}
	
	public static void PrintTables(Workbook workbook, ExternalConnection ec)
	{
		//Iterate all the worksheets
		for (int j = 0; j < workbook.getWorksheets().getCount(); j++)
		{
			Worksheet worksheet = workbook.getWorksheets().get(j);

			//Check all the query tables in a worksheet
			for (int k = 0; k < worksheet.getQueryTables().getCount(); k++)
			{
				QueryTable qt = worksheet.getQueryTables().get(k);

				//Check if query table is related to this external connection
				if (ec.getId() == qt.getConnectionId()
					&& qt.getConnectionId() >= 0)
				{
					//Print the query table name and print its "Refers To" range
					System.out.println("querytable " + qt.getName());
					String n = qt.getName().replace('+', '_').replace('=', '_');
					Name name = workbook.getWorksheets().getNames().get("'" + worksheet.getName() + "'!" + n);
					if (name != null)
					{
						Range range = name.getRange();
						if (range != null)
						{
							System.out.println("Refers To: " + range.getRefersTo());
						}
					}
				}
			}

			//Iterate all the list objects in this worksheet
			for (int k = 0; k < worksheet.getListObjects().getCount(); k++)
			{
				ListObject table = worksheet.getListObjects().get(k);

				//Check the data source type if it is query table
				if (table.getDataSourceType() == TableDataSourceType.QUERY_TABLE)
				{
					//Access the query table related to list object
					QueryTable qt = table.getQueryTable();

					//Check if query table is related to this external connection
					if (ec.getId() == qt.getConnectionId()
					&& qt.getConnectionId() >= 0)
					{
						//Print the query table name and print its refersto range
						System.out.println("querytable " + qt.getName());
						System.out.println("Table " + table.getDisplayName());
						System.out.println("refersto: " + worksheet.getName() + "!" + CellsHelper.cellIndexToName(table.getStartRow(), table.getStartColumn()) + ":" + CellsHelper.cellIndexToName(table.getEndRow(), table.getEndColumn()));
					}
				}
			}
		}
	}//end-PrintTables
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_find_reference_cells_from_external_connection);

	
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			System.out.println("---------------Executing--------------------------");
			findReferenceCellsFromExternalConnection();
			tx.setText("Successfully Done. Please see the console output of the code in the LogCat.");
			System.out.println("---------------Done--------------------------");
		}
		catch(Exception ex)
		{
			tx.setText("Error during document processing: " + ex.getMessage());
		}

	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(
				R.menu.find_reference_cells_from_external_connection, menu);
		return true;
	}

}
