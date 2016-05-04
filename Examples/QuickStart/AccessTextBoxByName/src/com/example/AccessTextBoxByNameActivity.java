package com.example;

import com.aspose.cells.*;

import android.os.Bundle;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;


public class AccessTextBoxByNameActivity extends Activity {
	
	public void accessTextBoxByName() throws Exception
	{
		Workbook workbook = new Workbook();

		Worksheet sheet = workbook.getWorksheets().get(0);

		int idx = sheet.getTextBoxes().add(10, 10, 10, 10);

		//Create a texbox with some text and assign it some name
		TextBox tb1 = sheet.getTextBoxes().get(idx);
		tb1.setName("MyTextBox");
		tb1.setText("This is MyTextBox");

		//Access the same textbox via its name
		TextBox tb2 = sheet.getTextBoxes().get("MyTextBox");

		//Displaying the text of the textbox accessed by its name
		System.out.println(tb2.getText());
	}

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_access_text_box_by_name);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);

		try
		{
			System.out.println("---------------Executing--------------------------");
			accessTextBoxByName();
			tx.setText("Successfully Done. Please see the console output of the code in the LogCat.");
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
		getMenuInflater().inflate(R.menu.access_text_box_by_name, menu);
		return true;
	}

}
