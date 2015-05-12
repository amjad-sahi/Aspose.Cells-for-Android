package com.example.datafilter;

import java.io.File;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void dataFilter() throws Exception {

		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		Workbook workbook = new Workbook();
	
		WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(0);

        Cells cells = worksheet.getCells();
        Cell cell;

        //Put a value into a cell
        cell = cells.get("A1");
        cell.setValue("Fruit");
        cell = cells.get("B1");
        cell.setValue("Total");
        cell = cells.get("A2");
        cell.setValue("Apple");
        cell = cells.get("B2");
        cell.setValue(1000);
        cell = cells.get("A3");
        cell.setValue("Orange");
        cell = cells.get("B3");
        cell.setValue(2500);
        cell = cells.get("A4");
        cell.setValue("Bananas");
        cell = cells.get("B4");
        cell.setValue(2500);
        cell = cells.get("A5");
        cell.setValue("Pear");
        cell = cells.get("B5");
        cell.setValue(1000);
        cell = cells.get("A6");
        cell.setValue("Grape");
        cell = cells.get("B6");
        cell.setValue(2000);

        cell = cells.get("D1");
        cell.setValue("Count:");
        cell = cells.get("E1");
        cell.setFormula("=SUBTOTAL(2, B1:B6)");

        worksheet.getAutoFilter().setRange("A1:B6");
        
        workbook.save(sdPath + "DataFilter.xlsx", SaveFormat.XLSX);
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			dataFilter();
			tx.setText("Data Filter created successfully. Please check the root of SD path.");
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
