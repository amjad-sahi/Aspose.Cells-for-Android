package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.util.Log;

import com.aspose.cells.Workbook;

import java.util.Random;

public class ReadCellValuesInMultipleThreads {

    private static final String TAG = ReadCellValuesInMultipleThreads.class.getName();

    public class ThreadProc implements Runnable {

        boolean isRunning = true;
        Workbook testWorkbook;
        Random r = new Random();

        public ThreadProc(Workbook workbook) {
            this.testWorkbook = workbook;
        }

        public int randomNext(int Low, int High) {
            int R = r.nextInt(High - Low) + Low;
            return R;
        }

        public void kill() {
            this.isRunning = false;
        }

        public void run() {

            while (this.isRunning) {
                int row = randomNext(0, 100);
                int col = randomNext(0, 10);

                String s = testWorkbook.getWorksheets().get(0).getCells().get(row, col).getStringValue();

                if (s.equals("R" + row + "C" + col) != true) {
                    System.out.println("This message box will show up when cells read values are incorrect.");
                }
            }
        }
    }

    public void readCellValuesInMultipleThreads() {
        try {
            Workbook testWorkbook = new Workbook();
            testWorkbook.getWorksheets().clear();
            testWorkbook.getWorksheets().add("Sheet1");

            for (int row = 0; row < 100; row++)
                for (int col = 0; col < 10; col++)
                    testWorkbook.getWorksheets().get(0).getCells().get(row, col).setValue("R" + row + "C" + col);

            //Commenting this line will show a pop-up message
            testWorkbook.getWorksheets().get(0).getCells().setMultiThreadReading(true);

            ThreadProc tp = new ThreadProc(testWorkbook);

            Thread myThread1 = new Thread(tp);
            myThread1.start();

            Thread myThread2 = new Thread(tp);
            myThread2.start();

            Thread.currentThread().sleep(5 * 1000);
            tp.kill();
        } catch (Exception e) {
            Log.e(TAG, "Reading Cells Values in Multiple Threads Simultaneously", e);
        }
    }
}
