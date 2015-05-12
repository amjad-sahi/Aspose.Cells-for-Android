package com.example.ungroupingrowsandcolumns;

import java.io.File;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void groupingRowsAndColumns() throws Exception {
		
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		String filePath = sdPath + "UnGroupingRowsAndColumns.xls";

        Workbook wb = new Workbook(filePath);
	    Worksheet worksheet = wb.getWorksheets().get(0);
	    
        Cells cells = worksheet.getCells();
        cells.ungroupRows(0, 9);
        cells.ungroupColumns(0, 1);
        
        wb.save(filePath + ".out.xls");
	
	}

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);

		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			groupingRowsAndColumns();
			tx.setText("UnGroupingRowsAndColumns created successfully. Please check the root of SD path.");
		}
		catch(Exception ex)
		{
			tx.setText("Error during document processing: " + ex.getMessage());
		}
		
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

}
