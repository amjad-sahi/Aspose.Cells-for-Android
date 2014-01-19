package com.example.adjustingrowsandcolumns;

import java.io.File;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void adjustingRowsAndColumns() throws Exception {

		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		int fileFormatType = FileFormatType.EXCEL_97_TO_2003;
	    Workbook workbook = new Workbook(fileFormatType);
	    WorksheetCollection worksheets = workbook.getWorksheets();
	    Worksheet worksheet = worksheets.get(0);
	
	    Cells cells = worksheet.getCells();
	
	    //Set the height of all row in the worksheet
	    cells.setStandardHeight(20);
	    //Set the width of all columns in the worksheet
	    cells.setStandardWidth(20);
	
	    //Set the width of the first column
	    cells.setColumnWidth(0, 12);
	    //Set the width of the second column
	    cells.setColumnWidth(1, 40);
	    //Setting the height of row
	    cells.setRowHeight(1, 8);
	    
		workbook.save(sdPath + "Cells_AdjustingRowsAndColumns.xls", SaveFormat.EXCEL_97_TO_2003);

	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			adjustingRowsAndColumns();
			tx.setText("AdjustingRowsAndColumns created successfully. Please check the root of SD path.");
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
