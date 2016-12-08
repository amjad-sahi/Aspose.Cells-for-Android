package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Scenario;
import com.aspose.cells.ScenarioInputCellCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CreateManipulateOrRemoveScenariosFromWorksheet {

    private static final String TAG = CreateManipulateOrRemoveScenariosFromWorksheet.class.getName();

    public void createManipulateOrRemoveScenarios() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load an Excel file
            Workbook workbook = new Workbook(filePath + "sample.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Remove the existing first scenario from the sheet
            worksheet.getScenarios().removeAt(0);

            //Create a scenario
            int i = worksheet.getScenarios().add("MyScenario");
            //Get the scenario
            Scenario scenario = worksheet.getScenarios().get(i);
            //Add comment to it
            scenario.setComment("Test sceanrio is created.");
            //Get the input cells for the scenario
            ScenarioInputCellCollection sic = scenario.getInputCells();
            //Add the scenario on B4 (as changing cell) with default value
            sic.add(3, 1, "1100000");

            //Save the Excel file.
            workbook.save(filePath + "Scenarios_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Create, Manipulate or Remove Scenarios from Worksheets", e);
        }
    }

}
