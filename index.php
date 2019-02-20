<?php
require __DIR__ . '/vendor/autoload.php';
Dompdf\Autoloader::register();
// reference the Dompdf namespace
use Dompdf\Dompdf;
use Dompdf\Options;

/**  Define a Read Filter class implementing \PhpOffice\PhpSpreadsheet\Reader\IReadFilter  */
class MyReadFilter implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter
{
    private $startRow = 0;
    private $endRow   = 0;
    private $columns  = [];

    /**  Get the list of rows and columns to read  */
    public function __construct($startRow, $endRow, $columns) {
        $this->startRow = $startRow;
        $this->endRow   = $endRow;
        $this->columns  = $columns;
    }

    public function readCell($column, $row, $worksheetName = '') {
        //  Only read the rows and columns that were configured
        if ($row >= $this->startRow && $row <= $this->endRow) {
            if (in_array($column,$this->columns)) {
                return true;
            }
        }
        return false;
    }
}

class PDFHelper {
    private $fileName;
    private $minPageNum = 1;
    private $maxPageNum = 20;
    private $worksheet;
    private $spreadsheet;

    public function getWorksheet($fileName, $download=false){
        /**  Create an Instance of our Read Filter  **/
        $filterSubset = new MyReadFilter($this->minPageNum, $this->maxPageNum,range('A','Y'));

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        // $reader->setLoadSheetsOnly("Sheet1");
        if ($download == false){
            $reader->setReadFilter($filterSubset);
        }
        $this->fileName = $fileName;
        $this->spreadsheet = $reader->load($fileName);
        $this->worksheet = $this->spreadsheet->getSheet(0)->toArray(null, true, true, true);

        return $this->worksheet;
    }

    public function getTable($download=false){
        //html table with the excel file data
        $this->getWorksheet($this->fileName, $download);
        $html_tb ='<font size="1" face="Courier New" >';
        $html_tb .='<table class="table" border="1"><tr><th>'. implode('</th><th>', $this->worksheet[1]) .'</th></tr>';
        $nr = count($this->worksheet); //number of rows
        for($i=2; $i<=$nr; $i++){
            $html_tb .='<tr><td>'. implode('</td><td>', $this->worksheet[$i]) .'</td></tr>';
        }
        $html_tb .='</table></font>';

        return $html_tb;
    }

    public function getHtml(){
        // added you custome html here to render the table
        // it will also have the button and pagination
        $html = '
        //combo box to show a number of records per npage
        <div>
        <form action="index.php" method="post">
        <label for="show_entries"> Show Entries : </label>
        <select id="show_entries" name="show_entries" >
           <option value="1">10</option>
           <option value="2">20</option>
           <option value="3">50</option>
           <option value="4">100</option>
        </select>

        <input type="submit" name="search" value="Search"/>
        </form>';

        $html .= $this->getTable();
        $html .=    '<form >
                <a href="?file=csv">
                <button type="button" onclick="downloadCSV()"> Download CSV </button>
                </a>
                <a href="?file=pdf">
                <button type="button" onclick="downloadPDF()"> Download PDF </button>
                </a>
                
            </form></div>';
        return $html;
    }

    function getCSV(){
        // redirect output to client browser
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment;filename="myfile.csv"');
        header('Cache-Control: max-age=0');

        $this->getWorksheet($this.$this->fileName, true);
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($this->spreadsheet);
        $writer->save('php://output');
    }

    function getPDF(){
        // instantiate and use the dompdf class
        $dompdf = new Dompdf();
        $dompdf->loadHtml($this->getTable(true));

        // (Optional) Setup the paper size and orientation
        $dompdf->setPaper('A2', 'landscape');

        // Render the HTML as PDF
        $dompdf->render();

        // Output the generated PDF to Browser
        $dompdf->stream();
    }
}


$fileName = __DIR__ ."/GooglePlaySept2018.xlsx";

$pdfHelper = new PDFHelper();
$pdfHelper->getWorksheet($fileName);

if (isset($_GET['file'])){
    if ($_GET['file'] == 'pdf') {
        echo $pdfHelper->getPDF();
    }elseif ($_GET['file'] == 'csv'){
        echo $pdfHelper->getCSV();
    }
}
//echo $_SERVER['minPage'];
//echo $_SERVER['maxPage'];
echo $pdfHelper->getHtml();