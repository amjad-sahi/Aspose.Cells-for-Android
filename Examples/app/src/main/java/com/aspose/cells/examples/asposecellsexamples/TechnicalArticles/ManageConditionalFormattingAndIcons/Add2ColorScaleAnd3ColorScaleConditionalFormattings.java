package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageConditionalFormattingAndIcons;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.Color;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class Add2ColorScaleAnd3ColorScaleConditionalFormattings {
    private static final String TAG = Add2ColorScaleAnd3ColorScaleConditionalFormattings.class.getName();

    public void add2ColorScaleAnd3ColorScaleConditionalFormattings() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook
            Workbook workbook = new Workbook();

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Add some data in cells
            worksheet.getCells().get("A1").putValue("2-Color Scale");
            worksheet.getCells().get("D1").putValue("3-Color Scale");

            for(int i=2; i<=15; i++) {
                worksheet.getCells().get("A" + i).putValue(i);
                worksheet.getCells().get("D" + i).putValue(i);
            }

            //Adding 2-Color Scale Conditional Formatting
            CellArea ca = CellArea.createCellArea("A2", "A15");

            int idx = worksheet.getConditionalFormattings().add();
            FormatConditionCollection fcc = worksheet.getConditionalFormattings().get(idx);
            fcc.addCondition(FormatConditionType.COLOR_SCALE);
            fcc.addArea(ca);

            FormatCondition fc = worksheet.getConditionalFormattings().get(idx).get(0);
            fc.getColorScale().setIs3ColorScale(false);
            fc.getColorScale().setMaxColor(Color.getLightBlue());
            fc.getColorScale().setMinColor(Color.getLightGreen());

            //Adding 3-Color Scale Conditional Formatting
            ca = CellArea.createCellArea("D2", "D15");

            idx = worksheet.getConditionalFormattings().add();
            fcc = worksheet.getConditionalFormattings().get(idx);
            fcc.addCondition(FormatConditionType.COLOR_SCALE);
            fcc.addArea(ca);

            fc = worksheet.getConditionalFormattings().get(idx).get(0);
            fc.getColorScale().setIs3ColorScale(true);
            fc.getColorScale().setMaxColor(Color.getLightBlue());
            fc.getColorScale().setMidColor(Color.getYellow());
            fc.getColorScale().setMinColor(Color.getLightGreen());

            //Save the workbook
            workbook.save(filePath +"Add2ColorScaleAnd3ColorScaleConditionalFormattings_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Adding 2-Color Scale and 3-Color Scale Conditional Formattings", e);
        }
    }
}
