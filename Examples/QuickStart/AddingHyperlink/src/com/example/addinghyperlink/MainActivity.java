package com.example.addinghyperlink;

import java.io.File;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void addingHyperlink() throws Exception {

		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		Workbook workbook = new Workbook();

		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		Cells cells = worksheet.getCells();
		Cell cell;
		/* Adding link to a URL */
		// Put a value into a cell
		cell = cells.get("A1");
		cell.setValue("Visit Aspose");
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setColor(Color.getBlue());
		font.setUnderline(FontUnderlineType.SINGLE);
		cell.setStyle(style);

		// Add a hyperlink to a URL at "A1" cell
		worksheet.getHyperlinks().add("A1", "B1", "http://www.aspose.com",
				"Hello Aspose", "1");
		
		workbook.save(sdPath + "AddingHyperlink.xlsx", SaveFormat.XLSX);
	}

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			addingHyperlink();
			tx.setText("Hyperlink created successfully. Please check the root of SD path.");
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
