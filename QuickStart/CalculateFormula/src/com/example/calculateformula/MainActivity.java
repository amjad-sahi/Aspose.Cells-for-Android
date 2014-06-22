package com.example.calculateformula;

import java.io.File;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void runCalculateFormula() throws Exception {

		// Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath()
				+ File.separator;

		// First push CalculateFormula.xls in SD card manually
		// Now open CalculateFormula.xls from SD card
		String filePath = sdPath + "CalculateFormula.xls";
		Workbook workbook = new Workbook(filePath);

		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);
		Cells cells = worksheet.getCells();
		Cell cell;
		String strFormula;

		//Read string formulas from column C and set formulas in column D
		for (int i = 7; i <= 131; i++) {
			cell = cells.get(i, 2);
			strFormula = cell.getStringValue();
			cell = cells.get(i, 3);
			cell.setFormula(strFormula);
		}

		//Calculate the formulas
		workbook.calculateFormula();
		
		//Read calculated values from column D and write them in column E
		for (int j = 7; j <= 131; j++) {
			cell = cells.get(j, 3);
			cells.get(j, 4).setValue(cell.getValue());
		}
		
		//Write the shared formula
		cells.get("F8").setSharedFormula("=E8=D8", 125, 1);

		// Save the output workbook in SD card
		workbook.save(filePath + ".out.xls");

	}

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);

		final TextView tx = (TextView) findViewById(R.id.myTextBox);

		try {
			runCalculateFormula();
			tx.setText("CalcaluateFormula file has been created successfully. Please check the root of SD path.");
		} catch (Exception ex) {
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
