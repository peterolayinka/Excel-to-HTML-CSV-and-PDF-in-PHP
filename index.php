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
        //  Read rows 1 to 7 and columns A to E only
        // if ($row >= 7 && $row <= 20) {
        //     if (in_array($column,range('A','E'))) {
        //         return true;
        //     }
        // }
        return false;
    }
}

class PDFHelper {
    private $fileName;
    private $paginator = 200;
    public $minPageNum = 0;
    public $maxPageNum;
    public $totalPages;
    public $totalRow;
    public $currentPage;
    public $nextPage;
    public $prevPage;
    private $worksheet;
    private $spreadsheet;

    public function getWorksheet($fileName, $download=false){
        /**  Create an Instance of our Read Filter  **/
        $filterSubset = new MyReadFilter($this->minPageNum, $this->maxPageNum,range('A','Y'));

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $unfilteredReader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setLoadSheetsOnly("Sheet1");
        $unfilteredReader->setLoadSheetsOnly("Sheet1");
        $this->totalRow = $reader->load($fileName)->getActiveSheet()->getHighestDataRow();
        $this->totalPages = round($this->totalRow / $this->paginator);
        if ($download == false){
            $reader->setReadFilter($filterSubset);
        }
        $this->fileName = $fileName;
        $this->spreadsheet = $reader->load($fileName);
        $this->unfilteredSpreadsheet = $unfilteredReader->load($fileName);
        $this->worksheet = $this->spreadsheet->getSheet(0)->toArray(null, true, true, true);
        $this->unfilteredWorksheet = $this->unfilteredSpreadsheet->getSheet(0)->toArray(null, true, true, true);

        return $this->worksheet;
    }

    public function getTable($download=false){
        //html table with the excel file data
        $this->getWorksheet($this->fileName, $download);
        // print_r ($this->worksheet);
        // $nr = count(array_filter($this->worksheet)); //number of rows
        // echo $nr.'Opps'. $this->totalRow;
        $html_tb ='';
        $html_tb .='<table class="table" style="font-size:10px;" border="1"><tr><th>SN</th><th>'. implode('</th><th>', $this->unfilteredWorksheet[1]) .'</th></tr>';
        // $html_tb .='<table class="table" style="font-size:10px;" border="1">';
        $startNum = ($this->currentPage == 1?$startNum = 2:$this->minPageNum);
        for($i=$startNum; $i<=$this->maxPageNum; $i++){
            if ($i > $this->totalRow){
                break;
            }
            $html_tb .='<tr><td>'.$i.'</td><td>'. implode('</td><td>', $this->worksheet[$i]) .'</td></tr>';
        }
        $html_tb .='</table>';

        return $html_tb;
    }

    public function getHtml(){
        // added you custome html here to render the table
        // it will also have the button and pagination
        //combo box to show a number of records per npage
        $html = '
        <div>' .
        'Total Pages: '. $this->totalPages.
        ' ------ Total Number of Rows: '. $this->totalRow.
        ' ------ Current Page: '. $this->currentPage.
        ' ------ Current Viewing Page: '. $this->minPageNum .' - '. $this->maxPageNum.
        '</br></br>';

        $html .= $this->getTable();
        if ($this->currentPage > 1){
            $html .= "<a href='?page={$this->prevPage}'><button>Prev</button></a> &nbsp; &nbsp;";
        }
        if ($this->currentPage < $this->totalPages){
            // $html .= '<a href=\"?pa÷ge='. $this->currentPage + 1 .'\"><button>Next</button></a> &nbsp; &nbsp;';
            $html .= "<a href='?page={$this->nextPage}'><button>Next</button></a> &nbsp; &nbsp;";
        }
        $html .= '
            <a href="?file=csv">
            <button type="button" onclick="downloadCSV()"> Download CSV </button>
            </a>

            <a href="?file=pdf&page='.$this->currentPage.'">
            <button type="button" onclick="downloadPDF()"> Download PDF </button>
            </a>

        </div>';
        return $html;
    }

    function getCSV(){
        // redirect output to client browser
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment;filename="myfile.csv"');
        header('Cache-Control: max-age=0');

        $this->getWorksheet($this->fileName, true);
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($this->spreadsheet);
        $writer->save('php://output');
    }

    function loadPageDependency($fileName){
        $this->getWorksheet($fileName);
        if (isset($_GET['page']) && $_GET['page'] != 0) {
            $this->currentPage = $_GET['page'];
        }else{
            $this->currentPage = 1;
        }

        $this->nextPage = (int)$this->currentPage + 1;
        $this->prevPage = (int)$this->currentPage - 1;
        $this->maxPageNum = $this->currentPage * $this->paginator;
        $this->minPageNum = (int)$this->prevPage * $this->paginator;
    }

    function getPDF(){
        // instantiate and use the dompdf class
        $dompdf = new Dompdf();
        $dompdf->loadHtml($this->getTable());

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

$pdfHelper->loadPageDependency($fileName);


if (isset($_GET['file'])){
    if ($_GET['file'] == 'pdf') {
        echo $pdfHelper->getPDF();
    }elseif ($_GET['file'] == 'csv'){
        echo $pdfHelper->getCSV();
    }
}else{
    echo $pdfHelper->getHtml();
}
// echo $_SERVER['minPage'];
// echo $_SERVER['maxPage'];