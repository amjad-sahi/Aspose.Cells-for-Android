package com.example.deleterowsandcolumns;

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

	void deleteRowsAndColumns() throws Exception {
		
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

        Workbook wb = new Workbook();
	    
        Worksheet worksheet = wb.getWorksheets().get(0);
        
        Cells cells = worksheet.getCells();
        
        //Put some values into cells
        Cell cell = cells.get("A1");
        cell.putValue("Row-1");

        cell = cells.get("A2");
        cell.putValue("Row-2");
        
        cell = cells.get("A3");
        cell.putValue("Row-3");
        
        cell = cells.get("A4");
        cell.putValue("Row-4");
        
        cell = cells.get("A5");
        cell.putValue("Row-5");
        
        cell = cells.get("B1");
        cell.putValue("Column B");
        
        cell = cells.get("C1");
        cell.putValue("Column C");
        
        cell = cells.get("D1");
        cell.putValue("Column D");
                
        //Delete 2 rows starting from 3rd row i.e 3rd and 4th rows
        cells.deleteRows(2, 2, false);
        
        //Delete 1 column starting from 2nd column i.e column B
        cells.deleteColumns(1, 1,false);
        
        wb.save(sdPath + "Cells_DeleteRowsAndColumns.xls");
        
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			deleteRowsAndColumns();
			tx.setText("DeletingRowsAndColumns created successfully. Please check the root of SD path.");
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
