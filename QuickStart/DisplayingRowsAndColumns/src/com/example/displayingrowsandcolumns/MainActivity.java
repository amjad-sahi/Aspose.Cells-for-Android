package com.example.displayingrowsandcolumns;

import java.io.File;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void displayingRowsAndColumns() throws Exception {
		
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

        Workbook wb = new Workbook();
	    
        Worksheet worksheet = wb.getWorksheets().get(0);
        
        //Display the 3rd row of the worksheet and set its height 25
        worksheet.getCells().unhideRow(2, 25);
        
        //Display the 2nd column of the worksheet and set its width 15
        worksheet.getCells().unhideColumn(1, 15);
        
        wb.save(sdPath + "Cells_DisplayRowsAndColumns.xls");
	
	}

	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			displayingRowsAndColumns();
			tx.setText("DisplayingRowsAndColumns created successfully. Please check the root of SD path.");
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
