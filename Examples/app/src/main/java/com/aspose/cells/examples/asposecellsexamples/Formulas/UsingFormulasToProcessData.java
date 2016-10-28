package com.aspose.cells.examples.asposecellsexamples.Formulas;

import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class UsingFormulasToProcessData {
    private static final String TAG = UsingFormulasToProcessData.class.getName();

    public void usingBuiltInFunctions() {
        try {
            Workbook workbook = new Workbook();
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Setting a complex formula on the 1st cell of the Cells collection of a worksheet
            worksheet.getCells().get(0).setFormula("=H7*(1+IF(P7 =$L$3,$M$3, (IF(P7=$L$4,$M$4,0))))");
        } catch (Exception e) {
            Log.e(TAG, "Using Built-in Functions", e);
        }
    }

    public void usingAddInFunctions() {
        try {
            Workbook workbook = new Workbook();
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Setting an add-in formula on the "H11" cell of the worksheet
            worksheet.getCells().get("H11").setAddInFormula("HRVSTTRK.xla", "=pct_overcut(F3:G3)");
        } catch (Exception e) {
            Log.e(TAG, "Using Add-in Functions", e);
        }
    }

    public void usingArrayFormula() {
        try {
            Workbook workbook = new Workbook();
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Setting an array formula on the cell "G2"
            worksheet.getCells().get("G2").setArrayFormula("=LINEST(E2:E12,A2:D12,TRUE,TRUE)", 5, 3);
        } catch (Exception e) {
            Log.e(TAG, "Using Array Formula", e);
        }
    }

    public void usingR1C1Formula() {
        try {
            Workbook workbook = new Workbook();
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Setting an R1C1 formula on the "A1" cell
            worksheet.getCells().get("A1").setR1C1Formula("=SUM(R[1]C[3]:R[3]C[4])");
        } catch (Exception e) {
            Log.e(TAG, "Using R1C1 Formula", e);
        }
    }
}
