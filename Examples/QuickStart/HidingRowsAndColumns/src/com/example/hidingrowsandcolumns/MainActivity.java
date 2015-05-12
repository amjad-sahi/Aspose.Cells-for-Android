package com.example.hidingrowsandcolumns;

import java.io.File;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void hidingRowsAndColumns() throws Exception {
		
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

        Workbook wb = new Workbook();
	    
        Worksheet worksheet = wb.getWorksheets().get(0);
        
        //Hide the 3rd row of the worksheet
        worksheet.getCells().hideRow(2);

        //Hide the 2nd column of the worksheet
        worksheet.getCells().hideColumn(1);
        
        wb.save(sdPath + "Cells_HideRowsAndColumns.xls");
	
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);

		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			hidingRowsAndColumns();
			tx.setText("HidingRowsAndColumns created successfully. Please check the root of SD path.");
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
