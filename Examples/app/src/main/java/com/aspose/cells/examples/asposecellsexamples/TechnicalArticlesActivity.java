package com.aspose.cells.examples.asposecellsexamples;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.ListView;
import android.widget.Toast;

import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ChangeDataSourceOfTheChartToDestinationWorksheet;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.DecreaseTheCalculationTimeOfCellCalculateMethod;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.DirectCalculationOfCustomFunction;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ExpandTextFromRightToLeftWhileExportingExcelFileToHTML;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ExportDataBarColorScaleAndIconSetConditionalFormatting;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.FilterDataWhileLoadingWorkbookFromTemplateFile;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.FindIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.GetOrSetTheClassIdentifierOfTheEmbeddedOLEObject;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat.ExportExcelToHTMLWithGridLines;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat.PreventExportingHiddenWorksheetContentsOnSavingToHTML;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat.SupportForLayoutOfDIVTags;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.AddActiveX;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.AutomaticallyRefreshOLEObject;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.ChangeCharacterSpacingOfExcelTextBoxOrShape;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.CreateTextBoxHavingEachLineWithDifferentHorizontalAlignment;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.SetCommentForTableOrListObject;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.SetLineSpacingOfTheParagraphInAShapeOrTextBox;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects.SetTextOfChartLegendEntryFillToNone;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections.ExternalDataConnectionOfTypeWebQuery;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections.FindQueryTablesAndListObjectsRelatedToExternalDataConnections;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections.FindReferenceCellsFromExternalConnection;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageExternalDataConnections.RetrieveExternalDataSourcesDetails;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart.RefreshAndCalculatePivotTableHavingCalculatedItems;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart.RestrictionsOfExcel2003WhileRefreshingPivotTable;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageVBAModules.AssignMacroCodeToFormControl;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageVBAModules.CheckIfVBACodeIsSigned;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageVBAModules.ModifyVBAOrMacroCode;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet.LoadSourceExcelFileWithoutCharts;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet.LoadWorkbookWithSpecifiedPrinterPaperSize;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ProvideExportedWorksheetHtmlFilePath;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.AddPDFBookmarks;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.CalculatePageSetupScalingFactor;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.ChangeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.CompressImagesForPDFConversion;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.ConfigureFontsForRenderingSpreadsheets;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.ConvertChartToImageInSVGFormat;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.CreateTransparentImageOfWorksheet;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.ExportRangeOfCellsInAWorksheetToImage;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.ExportWorksheetAndChartToSVGWithViewBoxAttribute;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.ExportWorksheetToImageByPage;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.FitAllWorksheetColumnsOnSinglePDFPage;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.GenerateAThumbnailImageOfAWorksheet;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.GetWarningsForFontSubstitutionWhileRenderingExcelFile;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.Implement1904DateSystem;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.LimitTheNumberOfPagesGeneratedExcelToPDFConversion;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.PageSetupAndPrintingOptions;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.RemoveWhiteSpacesFromTheDataBeforeRenderingToImage;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.SaveEachWorksheetToADifferentPDFFile;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.SaveExcelIntoPDFWithStandardOrMinimumSize;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint.SetDefaultFontWhileRenderingSpreadsheetToImage;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderUnicodeSupplementaryCharacters;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.SetShadowOfTextEffectsOfShapeOrTextBox;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.ApplySuperscriptAndSubscriptEffectsOnFonts;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.ConvertTextNumericDataToNumbersInAWorksheet;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.DisplayBulletsUsingHTML;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.Excel2007ThemesAndColors;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.ExtractThemeDataFromExcelFile;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.FormatWorksheetCellsInAWorkbook;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.GetCellStringValueWithAndWithoutFormatting;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.LineBreakAndTextWrapping;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.ModifyAnExistingStyle;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.RenderCustomDateFormatPattern;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.SpecifyCustomNumberDecimalAndGroupSeparatorsForWorkbook;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting.UsingBuiltInStyles;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.UpdateActiveXComboBoxControl;
import com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.UpdateReferencesInOtherWorksheets;

import java.util.ArrayList;

public class TechnicalArticlesActivity extends AppCompatActivity {

    private ListView listView;

    private AdapterView.OnItemClickListener sectionsListener = new AdapterView.OnItemClickListener() {
        @Override
        public void onItemClick(AdapterView<?> parent, View view, int position, long id) {

            if (ContextCompat.checkSelfPermission(TechnicalArticlesActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                Toast.makeText(TechnicalArticlesActivity.this, getString(R.string.write_to_external_storage_permission), Toast.LENGTH_SHORT).show();
                return;
            }

            switch (position) {
                case 0:
                    // Managing Workbooks & Worksheets
                    runManagingWorkbooksAndWorksheetsExamples();
                    break;
                case 1:
                    // Managing Rows, Columns & Cells
                    runManagingRowsColumnsAndCellsExamples();
                    break;
                case 2:
                    // Managing Charts, Shapes and Objects
                    runManagingChartsShapesAndObjectsExamples();
                    break;
                case 3:
                    // Managing Pivot Tables and Pivot Charts
                    runManagingPivotTablesAndPivotChartsExamples();
                    break;
                case 4:
                    // Managing Conditional Formatting & Icons
                    runManagingConditionalFormattingAndIconsExamples();
                    break;
                case 5:
                    // Managing External Data Connections
                    runManagingExternalDataConnectionsExamples();
                    break;
                case 6:
                    // Managing VBA Modules
                    runManagingVBAModulesExamples();
                    break;
                case 7:
                    // Rendering and Printing
                    runRenderingAndPrintingExamples();
                    break;
                case 8:
                    // Styling and Data Formatting
                    runStylingAndDataFormattingExamples();
                    break;
                case 9:
                    // Working with HTML Format
                    runWorkingWithHTMLFormatExamples();
                    break;
                case 10:
                    // Working with Calculation Engine
                    runWorkingWithCalculationEngineExamples();
                    break;
                case 11:
                    // Working with Smart Markers
                    runWorkingWithSmartMarkersExamples();
                    break;
                case 12:
                    // Miscellaneous
                    runMiscellaneousExamples();
                    break;
                default:
                    break;
            }
        }
    };

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_technical_articles);

        listView = (ListView)findViewById(R.id.technical_articles_lv);
        ArrayList<String> sectionsNames = new ArrayList<String>();
        sectionsNames.add("Managing Workbooks & Worksheets");
        sectionsNames.add("Managing Rows, Columns & Cells");
        sectionsNames.add("Managing Charts, Shapes and Objects");
        sectionsNames.add("Managing Pivot Tables and Pivot Charts");
        sectionsNames.add("Managing Conditional Formatting & Icons");
        sectionsNames.add("Managing External Data Connections");
        sectionsNames.add("Managing VBA Modules");
        sectionsNames.add("Rendering and Printing");
        sectionsNames.add("Styling and Data Formatting");
        sectionsNames.add("Working with HTML Format");
        sectionsNames.add("Working with Calculation Engine");
        sectionsNames.add("Working with Smart Markers");
        sectionsNames.add("Miscellaneous");

        ArrayAdapter<String> arrayAdapter = new ArrayAdapter<String>(this, android.R.layout.simple_list_item_1, sectionsNames);
        listView.setAdapter(arrayAdapter);

        listView.setOnItemClickListener(sectionsListener);
    }

    public void runManagingWorkbooksAndWorksheetsExamples() {

    }

    public void runManagingRowsColumnsAndCellsExamples() {

    }

    public void runManagingChartsShapesAndObjectsExamples() {

    }

    public void runManagingPivotTablesAndPivotChartsExamples() {

    }

    public void runManagingConditionalFormattingAndIconsExamples() {

    }

    public void runManagingExternalDataConnectionsExamples() {
        FindQueryTablesAndListObjectsRelatedToExternalDataConnections findQueryTables = new FindQueryTablesAndListObjectsRelatedToExternalDataConnections();
        findQueryTables.findQueryTablesAndListObjects();

        FindReferenceCellsFromExternalConnection findReferenceCells = new FindReferenceCellsFromExternalConnection();
        findReferenceCells.findReferenceCellsFromExternalConnection();

        RetrieveExternalDataSourcesDetails retrieveExternalDataSourcesDetails = new RetrieveExternalDataSourcesDetails();
        retrieveExternalDataSourcesDetails.retrieveExternalDataSourceDetails();

        ExternalDataConnectionOfTypeWebQuery externalDataConnectionOfTypeWebQuery = new ExternalDataConnectionOfTypeWebQuery();
        externalDataConnectionOfTypeWebQuery.workWithExternalDataConnectionOfTypeWebQuery();
    }

    public void runManagingVBAModulesExamples() {
        AssignMacroCodeToFormControl assignMacroCodeToFormControl = new AssignMacroCodeToFormControl();
        assignMacroCodeToFormControl.assignMacroCodeToFormControl();

        CheckIfVBACodeIsSigned checkIfVBACodeIsSigned = new CheckIfVBACodeIsSigned();
        checkIfVBACodeIsSigned.checkIfVBACodeIsSigned();

        ModifyVBAOrMacroCode modifyVBAOrMacroCode = new ModifyVBAOrMacroCode();
        modifyVBAOrMacroCode.modifyVBAOrMacroCode();
    }

    public void runRenderingAndPrintingExamples() {
        Implement1904DateSystem implement1904DateSystem = new Implement1904DateSystem();
        implement1904DateSystem.implement1904DateSystem();

        PageSetupAndPrintingOptions pageSetupAndPrintingOptions = new PageSetupAndPrintingOptions();
        pageSetupAndPrintingOptions.pageSetupOptions();
        pageSetupAndPrintingOptions.printOptions();

        LimitTheNumberOfPagesGeneratedExcelToPDFConversion limitTheNumberOfPagesGenerated = new LimitTheNumberOfPagesGeneratedExcelToPDFConversion();
        limitTheNumberOfPagesGenerated.limitTheNumberOfPagesGenerated();

        FitAllWorksheetColumnsOnSinglePDFPage fitAllWorksheetColumnsOnSinglePDFPage = new FitAllWorksheetColumnsOnSinglePDFPage();
        fitAllWorksheetColumnsOnSinglePDFPage.fitWorksheetColumnsOnSinglePDFPage();

        CreateTransparentImageOfWorksheet createTransparentImageOfWorksheet = new CreateTransparentImageOfWorksheet();
        createTransparentImageOfWorksheet.generateATransparentImage();

        GetWarningsForFontSubstitutionWhileRenderingExcelFile getWarningsForFontSubstitution = new GetWarningsForFontSubstitutionWhileRenderingExcelFile();
        getWarningsForFontSubstitution.getWarningsForFontSubstitution();

        CalculatePageSetupScalingFactor calculatePageSetupScalingFactor = new CalculatePageSetupScalingFactor();
        calculatePageSetupScalingFactor.calculatePageSetupScalingFactor();

        AddPDFBookmarks addPDFBookmarks = new AddPDFBookmarks();
        addPDFBookmarks.addPDFBookmarks();

        GenerateAThumbnailImageOfAWorksheet generateAThumbnailImageOfAWorksheet = new GenerateAThumbnailImageOfAWorksheet();
        generateAThumbnailImageOfAWorksheet.generateAThumbnailImageOfAWorksheet();

        SaveEachWorksheetToADifferentPDFFile saveEachWorksheetToADifferentPDFFile = new SaveEachWorksheetToADifferentPDFFile();
        saveEachWorksheetToADifferentPDFFile.saveEachWorksheetToADifferentPDFFile();

        ExportRangeOfCellsInAWorksheetToImage exportRangeOfCellsInAWorksheetToImage = new ExportRangeOfCellsInAWorksheetToImage();
        exportRangeOfCellsInAWorksheetToImage.exportRangeOfCellsInAWorksheetToImage();

        ChangeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF fontOfSpecificUnicodeCharactersWhileSavingToPDF = new ChangeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF();
        fontOfSpecificUnicodeCharactersWhileSavingToPDF.changeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF();

        SaveExcelIntoPDFWithStandardOrMinimumSize saveExcelIntoPDF = new SaveExcelIntoPDFWithStandardOrMinimumSize();
        saveExcelIntoPDF.saveExcelIntoPDFWithStandardOrMinimumSize();

        ExportWorksheetAndChartToSVGWithViewBoxAttribute exportWorksheetAndChartToSVGWithViewBoxAttribute = new ExportWorksheetAndChartToSVGWithViewBoxAttribute();
        exportWorksheetAndChartToSVGWithViewBoxAttribute.exportWorksheetAndChartToSVGWithViewBoxAttribute();

        ExportWorksheetToImageByPage exportWorksheetToImageByPage = new ExportWorksheetToImageByPage();
        exportWorksheetToImageByPage.renderFirstPageOfWorksheetToJPEGFormat();
        exportWorksheetToImageByPage.renderAllWorksheetsPrintingPagesToSeparateImages();

        ConvertChartToImageInSVGFormat convertChartToImageInSVGFormat = new ConvertChartToImageInSVGFormat();
        convertChartToImageInSVGFormat.convertChartToImageInSVGFormat();

        CompressImagesForPDFConversion compressImagesForPDFConversion = new CompressImagesForPDFConversion();
        compressImagesForPDFConversion.compressImagesForPDFConversion();

        RemoveWhiteSpacesFromTheDataBeforeRenderingToImage removeWhiteSpacesFromTheDataBeforeRenderingToImage = new RemoveWhiteSpacesFromTheDataBeforeRenderingToImage();
        removeWhiteSpacesFromTheDataBeforeRenderingToImage.removeWhiteSpacesFromTheDataBeforeRenderingToImage();

        SetDefaultFontWhileRenderingSpreadsheetToImage setDefaultFontWhileRenderingSpreadsheetToImage = new SetDefaultFontWhileRenderingSpreadsheetToImage();
        setDefaultFontWhileRenderingSpreadsheetToImage.setDefaultFontWhileRenderingSpreadsheetToImages();

        ConfigureFontsForRenderingSpreadsheets configureFontsForRenderingSpreadsheets = new ConfigureFontsForRenderingSpreadsheets();
        configureFontsForRenderingSpreadsheets.fontSubstitutionMechanism();

    }

    public void runStylingAndDataFormattingExamples() {
        LineBreakAndTextWrapping lineBreakAndTextWrapping = new LineBreakAndTextWrapping();
        lineBreakAndTextWrapping.wrappingText();
        lineBreakAndTextWrapping.explicitLineBreaks();

        ModifyAnExistingStyle modifyAnExistingStyle = new ModifyAnExistingStyle();
        modifyAnExistingStyle.createAndModifyAStyle();
        modifyAnExistingStyle.modifyAStyleInATemplateFile();

        ExtractThemeDataFromExcelFile extractThemeDataFromExcelFile = new ExtractThemeDataFromExcelFile();
        extractThemeDataFromExcelFile.extractThemeDataFromExcelFile();

        GetCellStringValueWithAndWithoutFormatting getCellStringValueWithAndWithoutFormatting = new GetCellStringValueWithAndWithoutFormatting();
        getCellStringValueWithAndWithoutFormatting.getCellStringValueWithAndWithoutFormatting();

        RenderCustomDateFormatPattern renderCustomDateFormatPattern = new RenderCustomDateFormatPattern();
        renderCustomDateFormatPattern.renderCustomDateFormatPattern();

        ConvertTextNumericDataToNumbersInAWorksheet convertTextNumericDataToNumbersInAWorksheet = new ConvertTextNumericDataToNumbersInAWorksheet();
        convertTextNumericDataToNumbersInAWorksheet.convertTextNumericDataToNumbersInAWorksheet();

        FormatWorksheetCellsInAWorkbook formatWorksheetCellsInAWorkbook = new FormatWorksheetCellsInAWorkbook();
        formatWorksheetCellsInAWorkbook.formatWorksheetCellsInAWorkbook();

        Excel2007ThemesAndColors excel2007ThemesAndColors = new Excel2007ThemesAndColors();
        excel2007ThemesAndColors.getAndSetThemeColors();
        excel2007ThemesAndColors.applyCustomThemes();
        excel2007ThemesAndColors.usingThemeColors();

        DisplayBulletsUsingHTML displayBulletsUsingHTML = new DisplayBulletsUsingHTML();
        displayBulletsUsingHTML.displayBulletsUsingHTML();

        UsingBuiltInStyles builtInStyles = new UsingBuiltInStyles();
        builtInStyles.usingBuiltInStyles();

        ApplySuperscriptAndSubscriptEffectsOnFonts superscriptAndSubscriptEffectsOnFonts = new ApplySuperscriptAndSubscriptEffectsOnFonts();
        superscriptAndSubscriptEffectsOnFonts.setSuperscriptEffect();
        superscriptAndSubscriptEffectsOnFonts.setSubscriptEffect();

        SpecifyCustomNumberDecimalAndGroupSeparatorsForWorkbook customNumberDecimalAndGroupSeparatorsForWorkbook = new SpecifyCustomNumberDecimalAndGroupSeparatorsForWorkbook();
        customNumberDecimalAndGroupSeparatorsForWorkbook.specifyCustomSeparators();

    }

    public void runWorkingWithHTMLFormatExamples() {
        PreventExportingHiddenWorksheetContentsOnSavingToHTML preventExportingHiddenWorksheetContentsOnSavingToHTML =
                new PreventExportingHiddenWorksheetContentsOnSavingToHTML();
        preventExportingHiddenWorksheetContentsOnSavingToHTML.preventExportingHiddenWorksheetContentsOnSavingToHTML();
    }

    public void runWorkingWithCalculationEngineExamples() {

    }

    public void runWorkingWithSmartMarkersExamples() {

    }

    public void runMiscellaneousExamples() {
        // Add Toggle Button ActiveX Control
        AddActiveX addActiveX = new AddActiveX();
        addActiveX.addToggleButtonActiveXControl();

        // Filtering the kind of data while loading the workbook from template file
        FilterDataWhileLoadingWorkbookFromTemplateFile filterData = new FilterDataWhileLoadingWorkbookFromTemplateFile();
        filterData.filterDataWhileLoadingWorkbookFromTemplateFile();

        // Update references in other worksheets while deleting blank columns and rows in a worksheet
        UpdateReferencesInOtherWorksheets updateReferencesInOtherWorksheets = new UpdateReferencesInOtherWorksheets();
        updateReferencesInOtherWorksheets.updateReferencesInOtherWorksheetsWhileDeletingBlankColumnsAndRowsInAWorksheet();

        // Export Excel to HTML with GridLines
        ExportExcelToHTMLWithGridLines exportExcel = new ExportExcelToHTMLWithGridLines();
        exportExcel.exportExcelToHTMLWithGridLines();

        // Automatically Refresh OLE Object via Microsoft Excel using Aspose.Cells
        AutomaticallyRefreshOLEObject automaticallyRefreshOLEObject = new AutomaticallyRefreshOLEObject();
        automaticallyRefreshOLEObject.automaticallyRefreshOLEObject();

        // Set the Comment for Table or List Object inside the Worksheet
        SetCommentForTableOrListObject setComment = new SetCommentForTableOrListObject();
        setComment.setCommentForTableOrListObject();

        // Set Default Font while Rendering Spreadsheet to Images
        SetDefaultFontWhileRenderingSpreadsheetToImage setDefaultFont = new SetDefaultFontWhileRenderingSpreadsheetToImage();
        setDefaultFont.setDefaultFontWhileRenderingSpreadsheetToImages();

        // Restrictions of Excel 2003 while Refreshing Pivot Table
        RestrictionsOfExcel2003WhileRefreshingPivotTable restrictionsOfExcel = new RestrictionsOfExcel2003WhileRefreshingPivotTable();
        restrictionsOfExcel.restrictionsOfExcel2003WhileRefreshingPivotTable();

        // Setting Shadow of Text Effects of Shape or TextBox
        SetShadowOfTextEffectsOfShapeOrTextBox shadowOfTextEffects = new SetShadowOfTextEffectsOfShapeOrTextBox();
        shadowOfTextEffects.setShadowOfTextEffectsOfShapeOrTextBox();

        // Direct calculation of custom function without writing it in a worksheet
        DirectCalculationOfCustomFunction directCalculation = new DirectCalculationOfCustomFunction();
        directCalculation.directCalculationOfCustomFunction();

        // Export DataBar, ColorScale and IconSet Conditional Formatting while Excel to HTML Conversion
        ExportDataBarColorScaleAndIconSetConditionalFormatting exportData = new ExportDataBarColorScaleAndIconSetConditionalFormatting();
        exportData.exportDataBarColorScaleAndIconSetConditionalFormatting();

        // Render Unicode Supplementary characters in output Pdf by Aspose.Cells
        RenderUnicodeSupplementaryCharacters renderUnicode = new RenderUnicodeSupplementaryCharacters();
        renderUnicode.renderUnicodeSupplementaryCharactersInOutputPdf();

        // Refresh and Calculate Pivot Table having Calculated Items
        RefreshAndCalculatePivotTableHavingCalculatedItems refreshAndCalculate = new RefreshAndCalculatePivotTableHavingCalculatedItems();
        refreshAndCalculate.refreshAndCalculatePivotTableHavingCalculatedItems();

        // Change Character Spacing of Excel TextBox or Shape
        ChangeCharacterSpacingOfExcelTextBoxOrShape changeCharacterSpacing = new ChangeCharacterSpacingOfExcelTextBoxOrShape();
        changeCharacterSpacing.changeCharacterSpacingOfExcelTextBoxOrShape();

        // Load Source Excel File Without Charts
        LoadSourceExcelFileWithoutCharts loadSourceExcelFile = new LoadSourceExcelFileWithoutCharts();
        loadSourceExcelFile.loadSourceExcelFileWithoutCharts();

        // Support for Layout of DIV Tags while Loading HTML
        SupportForLayoutOfDIVTags supportForLayout = new SupportForLayoutOfDIVTags();
        supportForLayout.supportForLayoutOfDIVTagsWhileLoadingHTML();

        // Set Text of Chart Legend Entry Fill to None
        SetTextOfChartLegendEntryFillToNone setText = new SetTextOfChartLegendEntryFillToNone();
        setText.setTextOfChartLegendEntryFillToNone();

        // Load Workbook with Specified Printer Paper Size
        LoadWorkbookWithSpecifiedPrinterPaperSize loadWorkbook = new LoadWorkbookWithSpecifiedPrinterPaperSize();
        loadWorkbook.loadWorkbookWithSpecifiedPrinterPaperSize();

        // Set Line Spacing of the Paragraph in a Shape or TextBox
        SetLineSpacingOfTheParagraphInAShapeOrTextBox lineSpacingOfTheParagraph = new SetLineSpacingOfTheParagraphInAShapeOrTextBox();
        lineSpacingOfTheParagraph.setLineSpacingOfTheParagraphInAShapeOrTextBox();

        // Create TextBox Having Each Line with Different Horizontal Alignment
        CreateTextBoxHavingEachLineWithDifferentHorizontalAlignment createTextBox = new CreateTextBoxHavingEachLineWithDifferentHorizontalAlignment();
        createTextBox.createTextBoxHavingEachLineWithDifferentHorizontalAlignment();

        // Change Data Source of the Chart to Destination Worksheet while Copying Rows or Range
        ChangeDataSourceOfTheChartToDestinationWorksheet changeDataSource = new ChangeDataSourceOfTheChartToDestinationWorksheet();
        changeDataSource.changeDataSourceOfTheChartToDestinationWorksheetWhileCopyingRowsOrRange();

        // Configuring Fonts for Rendering Spreadsheets
        ConfigureFontsForRenderingSpreadsheets fontsForRenderingSpreadsheets = new ConfigureFontsForRenderingSpreadsheets();
        fontsForRenderingSpreadsheets.selectionOfFonts();

        // Decrease the Calculation Time of Cell.Calculate() method
        DecreaseTheCalculationTimeOfCellCalculateMethod decreaseTheCalculationTime = new DecreaseTheCalculationTimeOfCellCalculateMethod();
        decreaseTheCalculationTime.decreaseTheCalculationTimeOfCellCalculateMethod();

        // Expanding text from right to left while exporting Excel file to HTML
        ExpandTextFromRightToLeftWhileExportingExcelFileToHTML expandText = new ExpandTextFromRightToLeftWhileExportingExcelFileToHTML();
        expandText.expandTextFromRightToLeftWhileExportingExcelFileToHTML();

        // Find if Data Points are in the Second Pie or Bar on a Pie of Pie or Bar of Pie Chart
        FindIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart findIfDataPoints = new FindIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart();
        findIfDataPoints.findIfDataPointsAreInTheSecondPieOrBarOnAPieOfPieOrBarOfPieChart();

        // Get or Set the Class Identifier of the Embedded OLE Object
        GetOrSetTheClassIdentifierOfTheEmbeddedOLEObject embeddedOLEObject = new GetOrSetTheClassIdentifierOfTheEmbeddedOLEObject();
        embeddedOLEObject.getOrSetTheClassIdentifierOfTheEmbeddedOLEObject();

        // Provide exported worksheet html file path via IFilePathProvider interface
        ProvideExportedWorksheetHtmlFilePath exportedWorksheetHtmlFilePath = new ProvideExportedWorksheetHtmlFilePath();
        exportedWorksheetHtmlFilePath.provideExportedWorksheetHtmlFilePath();

        // Update ActiveX ComboBox Control
        UpdateActiveXComboBoxControl comboBoxControl = new UpdateActiveXComboBoxControl();
        comboBoxControl.updateActiveXComboBoxControl();
    }

}
