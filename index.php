<?
require __DIR__ . '/vendor/autoload.php';
Dompdf\Autoloader::register();
// reference the Dompdf namespace
use Dompdf\Dompdf;

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

    /* set value to 'private $name' property */
    public function getWorksheet($fileName){
        /**  Create an Instance of our Read Filter  **/
        $filterSubset = new MyReadFilter($this->minPageNum, $this->maxPageNum,range('A','Y'));

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        // $reader->setLoadSheetsOnly("Sheet1");
        /**  Tell the Reader that we want to use the Read Filter  **/
        $reader->setReadFilter($filterSubset);
        $this->spreadsheet = $reader->load($fileName);
        $this->worksheet = $this->spreadsheet->getSheet(0)->toArray(null, true, true, true);

        return $this->worksheet;
    }

    public function getHTML(){
        // echo '<pre>';
        // print_r($this->worksheet);
        // echo '--------------';
        // // print_r($spreadsheet);
        // echo '</pre>';

        //now it is created a html table with the excel file data
        $html_tb ='<table border="1"><tr><th>'. implode('</th><th>', $this->worksheet[1]) .'</th></tr>';
        $nr = count($this->worksheet); //number of rows
        for($i=2; $i<=$nr; $i++){
          $html_tb .='<tr><td>'. implode('</td><td>', $this->worksheet[$i]) .'</td></tr>';
        }
        $html_tb .='</table>';
        return $html_tb;
    }

}


$fileName = __DIR__ ."/GooglePlaySept2018.xlsx";

$pdfHelper = new PDFHelper($fileName);
$pdfHelper->getWorksheet($fileName);
echo $pdfHelper->getHTML();






// echo $html_tb;




// echo '<pre>';
// var_dump($oTrends);
// var_dump($spreadsheet);
// print_r($worksheet);
// echo '--------------';
// // print_r($spreadsheet);
// echo '</pre>';
// // print $spreadsheet;

// $writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($spreadsheet);
// $writer->save("05featuredemo.csv");

// $writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($spreadsheet);

// $writer->save("05featuredemo.html");

// instantiate and use the dompdf class
// $dompdf = new Dompdf();
// $dompdf->loadHtml($html_tb);

// // (Optional) Setup the paper size and orientation
// $dompdf->setPaper('A4', 'landscape');

// // Render the HTML as PDF
// $dompdf->render();

// // Output the generated PDF to Browser
// $dompdf->stream();