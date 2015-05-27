<?php

use Aspose\Cells\CellsApi;

class CellsApiTest extends PHPUnit_Framework_TestCase {
    
    protected $cells;

    protected function setUp()
    {        
        $this->cells = new CellsApi();
    }
    
    public function testGetChartArea()
    {
        $result = $this->cells->GetChartArea($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetChartAreaBorder()
    {
        $result = $this->cells->GetChartAreaBorder($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetChartAreaFillFormat()
    {
        $result = $this->cells->GetChartAreaFillFormat($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetDocumentProperties()
    {
        $result = $this->cells->GetDocumentProperties($name="test_cells.xlsx", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetDocumentProperty()
    {
        $result = $this->cells->GetDocumentProperty($name="test_cells.xlsx", $propertyName="Author", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetExtractBarcodes()
    {
        $result = $this->cells->GetExtractBarcodes($name="test_cells.xlsx", $sheetName="Sheet1", $pictureNumber="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBook()
    {
        $result = $this->cells->GetWorkBook($name="test_cells.xlsx", $password = null, $isAutoFit = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookDefaultStyle()
    {
        $result = $this->cells->GetWorkBookDefaultStyle($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookName()
    {
        $result = $this->cells->GetWorkBookName($name="test_cells.xlsx", $nameName="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookNames()
    {
        $result = $this->cells->GetWorkBookNames($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookTextItems()
    {
        $result = $this->cells->GetWorkBookTextItems($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookWithFormat()
    {
        $result = $this->cells->GetWorkBookWithFormat($name="test_cells.xlsx", $format="pdf", $password = null, $isAutoFit = null, $storage = null, $folder = null, $outPath = null);
        $fh = fopen(getcwd(). '/Data/Output/Workbook.pdf', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Workbook.pdf');
    }
    
    public function testGetWorkSheet()
    {
        $result = $this->cells->GetWorkSheet($name="test_cells.xlsx", $sheetName="Sheet1", $verticalResolution = null, $horizontalResolution = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetCalculateFormula()
    {
        $result = $this->cells->GetWorkSheetCalculateFormula($name="test_cells.xlsx", $sheetName="Sheet3", $formula="SUM(A3,A4)", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetComment()
    {
        $result = $this->cells->GetWorkSheetComment($name="test_cells.xlsx", $sheetName="Sheet1", $cellName="A2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetComments()
    {
        $result = $this->cells->GetWorkSheetComments($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetHyperlink()
    {
        $result = $this->cells->GetWorkSheetHyperlink($name="test_cells.xlsx", $sheetName="Sheet3", $hyperlinkIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetHyperlinks()
    {
        $result = $this->cells->GetWorkSheetHyperlinks($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetMergedCell()
    {
        $result = $this->cells->GetWorkSheetMergedCell($name="test_cells.xlsx", $sheetName="Sheet3", $mergedCellIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetMergedCells()
    {
        $result = $this->cells->GetWorkSheetMergedCells($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetTextItems()
    {
        $result = $this->cells->GetWorkSheetTextItems($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetValidation()
    {
        $result = $this->cells->GetWorkSheetValidation($name="test_cells.xlsx", $sheetName="Sheet1", $validationIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetValidations()
    {
        $result = $this->cells->GetWorkSheetValidations($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetWithFormat()
    {
        $result = $this->cells->GetWorkSheetWithFormat($name="test_cells.xlsx", $sheetName="Sheet1", $format="png", $verticalResolution = null, $horizontalResolution = null, $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Sheet1.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Sheet1.png');
    }
    
    public function testGetWorkSheets()
    {
        $result = $this->cells->GetWorkSheets($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetAutoshape()
    {
        $result = $this->cells->GetWorksheetAutoshape($name="test_cells.xlsx", $sheetName="Sheet2", $autoshapeNumber=2, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetAutoshapeWithFormat()
    {
        $result = $this->cells->GetWorksheetAutoshapeWithFormat($name="test_cells.xlsx", $sheetName="Sheet2", $autoshapeNumber=2, $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Autoshape.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Autoshape.png');
    }
    
    public function testGetWorksheetAutoshapes()
    {
        $result = $this->cells->GetWorksheetAutoshapes($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetCell()
    {
        $result = $this->cells->GetWorksheetCell($name="test_cells.xlsx", $sheetName="Sheet1", $cellOrMethodName="A1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetCellStyle()
    {
        $result = $this->cells->GetWorksheetCellStyle($name="test_cells.xlsx", $sheetName="Sheet1", $cellName="A1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetCells()
    {
        $result = $this->cells->GetWorksheetCells($name="test_cells.xlsx", $sheetName="Sheet1", $offest = null, $count = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetChart()
    {
        $result = $this->cells->GetWorksheetChart($name="test_cells.xlsx", $sheetName="Sheet1", $chartNumber="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetChartLegend()
    {
        $result = $this->cells->GetWorksheetChartLegend($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetChartWithFormat()
    {
        $result = $this->cells->GetWorksheetChartWithFormat($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Chart.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Chart.png');
    }
    
    public function testGetWorksheetCharts()
    {
        $result = $this->cells->GetWorksheetCharts($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetColumn()
    {
        $result = $this->cells->GetWorksheetColumn($name="test_cells.xlsx", $sheetName="Sheet1", $columnIndex=1, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetColumns()
    {
        $result = $this->cells->GetWorksheetColumns($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetOleObject()
    {
        $result = $this->cells->GetWorksheetOleObject($name="test_cells.xlsx", $sheetName="Sheet2", $objectNumber="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetOleObjectWithFormat()
    {
        $result = $this->cells->GetWorksheetOleObjectWithFormat($name="test_cells.xlsx", $sheetName="Sheet2", $objectNumber="0", $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Ole.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Ole.png');
    }
    
    public function testGetWorksheetOleObjects()
    {
        $result = $this->cells->GetWorksheetOleObjects($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPicture()
    {
        $result = $this->cells->GetWorksheetPicture($name="test_cells.xlsx", $sheetName="Sheet2", $pictureNumber="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPictureWithFormat()
    {
        $result = $this->cells->GetWorksheetPictureWithFormat($name="test_cells.xlsx", $sheetName="Sheet2", $pictureNumber="0", $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Picture.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Picture.png');
    }
    
    public function testGetWorksheetPictures()
    {
        $result = $this->cells->GetWorksheetPictures($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPivotTable()
    {
        $result = $this->cells->GetWorksheetPivotTable($name="test_cells.xlsx", $sheetName="Sheet1", $pivottableIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPivotTables()
    {
        $result = $this->cells->GetWorksheetPivotTables($name="test_cells.xlsx", $sheetName="Sheet4", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetRow()
    {
        $result = $this->cells->GetWorksheetRow($name="test_cells.xlsx", $sheetName="Sheet2", $rowIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetRows()
    {
        $result = $this->cells->GetWorksheetRows($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
      
                         
}    