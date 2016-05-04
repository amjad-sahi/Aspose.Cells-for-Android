package com.example;

import java.util.Locale;

import com.aspose.cells.*;

import android.os.Bundle;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class ImplementCustomCalculationEngineActivity extends Activity {

	public class CustomEngine extends AbstractCalculationEngine
	{
		public void calculate(CalculationData data)
		{
			if(data.getFunctionName().toUpperCase().equals("SUM")==true)
			{
				Double val = (Double)data.getCalculatedValue();
				val = val + 30;
				
				data.setCalculatedValue(val);
			}
		}
	}
	
	public void implementCustomCalculationEngine() throws Exception
	{
		Workbook workbook = new Workbook();

		Worksheet sheet = workbook.getWorksheets().get(0);

		Cell a1 = sheet.getCells().get("A1");
		a1.setFormula("=Sum(B1:B2)");

		sheet.getCells().get("B1").putValue(10);
		sheet.getCells().get("B2").putValue(10);

		//Without Custom Engine, the value of cell A1 will be 20
		workbook.calculateFormula();

		System.out.println("Without Custom Engine Value of A1: " + a1.getStringValue());

		//With Custom Engine, the value of cell A1 will be 50
		CustomEngine engine = new CustomEngine();

		CalculationOptions opts = new CalculationOptions();
		opts.setCustomEngine(engine);

		workbook.calculateFormula(opts);

		System.out.println("With Custom Engine Value of A1: " + a1.getStringValue());
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_implement_custom_calculation_engine);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);

		try
		{
			System.out.println("---------------Executing--------------------------");
			implementCustomCalculationEngine();
			tx.setText("Successfully Done. Please see the console output of the code in LogCat.");
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
		getMenuInflater().inflate(R.menu.implement_custom_calculation_engine,
				menu);
		return true;
	}

}
