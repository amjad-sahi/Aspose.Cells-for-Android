package com.example;

import java.io.File;
import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

import com.aspose.cells.*;

public class DeleteRedundantSpacesFromHtmlActivity extends Activity {

	public void deleteRedundantSpacesFromHtml() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		//Sample Html containing redundant spaces after <br> tag
		String html = "<html>"
						+ "<body>"
							+ "<table>"
								+ "<tr>"
									+ "<td>"
										+ "<br>    This is sample data"
										+ "<br>    This is sample data"
										+ "<br>    This is sample data"
									+ "</td>"
								+ "</tr>"
							+ "</table>"
						+ "</body>"
					+ "</html>";

		//Convert Html to byte array
		byte[] byteArray = html.getBytes();

		//Set Html load options and keep precision true
		HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
		loadOptions.setDeleteRedundantSpaces(true);

		//Convert byte array into stream
		java.io.ByteArrayInputStream stream = new java.io.ByteArrayInputStream(byteArray);

		//Create workbook from stream with Html load options
		Workbook workbook = new Workbook(stream, loadOptions);

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Auto fit the sheet columns
		worksheet.autoFitColumns();

		//Save the workbook
		workbook.save(sdPath + "output-" + loadOptions.getDeleteRedundantSpaces() + ".xlsx", SaveFormat.XLSX);

	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_delete_redundant_spaces_from_html);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			System.out.println("---------------Executing--------------------------");
			deleteRedundantSpacesFromHtml();
			tx.setText("Documents created successfully. Please check the root of SD path.");
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
		getMenuInflater().inflate(R.menu.delete_redundant_spaces_from_html,
				menu);
		return true;
	}

}
