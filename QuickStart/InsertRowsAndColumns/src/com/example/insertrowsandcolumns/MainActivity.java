package com.example.insertrowsandcolumns;

import java.io.File;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void insertRowsAndColumns() throws Exception {
		
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

        Workbook wb = new Workbook();
	    
        Worksheet worksheet = wb.getWorksheets().get(0);
        
        Cells cells = worksheet.getCells();
        
        //Put some values into cells
        Cell cell = cells.get("A1");
        cell.putValue("Aspose");
        cell = cells.get("A2");
        cell.putValue(123);
        cell = cells.get("A3");
        cell.putValue("Hello World");
        cell = cells.get("B1");
        cell.putValue(120);
        
        //Insert a row or column into the worksheet
        
        //Insert 10 rows starting from 3rd row
        cells.insertRows(2, 10);
        
        //Insert 1 column starting from 2nd column 
        cells.insertColumns(1, 1);
               
        wb.save(sdPath + "Cells_InsertRowsAndColumns.xls");
	
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			insertRowsAndColumns();
			tx.setText("InsertingRowsAndColumns created successfully. Please check the root of SD path.");
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
