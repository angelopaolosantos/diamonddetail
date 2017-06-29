<?php  
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xls;

class DiamondDetail {
  public function __construct($post, $files){
    $this->post = $post;
    $this->file = $files;
  }

  public function getPost(){
    return $this->post;
  }

  public function getFiles(){
    return $this->file;
  }

  private function loadExcelFile($file){
    $this->reader = new Xls();
    $this->reader->setReadDataOnly(true);
    $this->spreadsheet = $this->reader->load($file);
  }

  private function loadTemplate($file){
    $this->reader = new Xls();
    $this->reader->setReadDataOnly(true);
    $this->spreadsheetTemplate = $this->reader->load($file);
  }

  private function RemoveExtension($strName) {    
    $ext = strrchr($strName, '.');    
    if($ext !== false) {       
      $strName = substr($strName, 0, -strlen($ext));    
    }
    return $strName; 
  }

  public function readExcelFile(){
    $this->loadExcelFile($this->file['file']['tmp_name']);
    $uniqueComponents=array();

    $currentRow=$this->post['start'];
    
    $sheet = $this->spreadsheet->getActiveSheet();

    while($currentRow < $this->post['end']){
      $invoice = $sheet->getCell('A'.$currentRow)->getCalculatedValue();

      $stylecode = $sheet->getCell('C'.$currentRow)->getCalculatedValue();

      $quantity = $sheet->getCell('D'.$currentRow)->getCalculatedValue();

      $reference = $sheet->getCell('E'.$currentRow)->getCalculatedValue();

      $item = array();

      $item['invoice'] = $invoice;
      $item['stylecode'] = $stylecode;
      $item['quantity'] = $quantity;
      $item['reference'] = $reference;

      //----Get Components
      
      $currentColumn = $this->post['component-start'];

      $components=array();

      while($currentColumn < $this->post['component-end']){

        $component=array();

        $componentName = $sheet->getCellByColumnAndRow($currentColumn, $currentRow)->getCalculatedValue();

        if($componentName){
          $componentQuantity = $sheet->getCellByColumnAndRow($currentColumn+2, $currentRow)->getCalculatedValue();
          $componentWeight = $sheet->getCellByColumnAndRow($currentColumn+3, $currentRow)->getCalculatedValue();

          $component['name']=$componentName;
          $component['quantity']=$componentQuantity;
          $component['weight']=$componentWeight;

          $components[]=$component;

          if (!in_array($componentName, $uniqueComponents)){
            $uniqueComponents[]=$componentName;
          }
        }
        $currentColumn+=5;
      }

      $item['components']=$components;

      //----End of Get Components
      $items[] = $item; 
      $currentRow++;

    }
    sort($uniqueComponents);

    $data['items']=$items;
    $data['unique']=$uniqueComponents;

    return $data;
  }

  public function writeExcelFile($data, $saveAsFilename=null, $template=null){
    if ($template==null){
      $spreadsheet = new Spreadsheet();
    }else{
      $this->loadTemplate($template);
      $spreadsheet = $this->spreadsheetTemplate;
    }

    if ($saveAsFilename==null){
      $newFilename=$this->RemoveExtension($this->file['file']['name']).'-result.xlsx';
      $saveAs='./files/results/'.$newFilename;
    }else{
      $saveAs='./files/results/'.$saveAsFilename;
    }

    $sheet = $spreadsheet->getActiveSheet();

    $items=$data['items'];
    $uniqueComponents=$data['unique'];
    
    $writeColumn=5;
    $writeRow=1;

    foreach($uniqueComponents as $uniqueComponent){
      $sheet->setCellValueByColumnAndRow($writeColumn, $writeRow, $uniqueComponent);
      $writeColumn++;
    }

    $writeRow++;

    foreach($items as $item){
      $sheet->setCellValue('A'.$writeRow, $item['invoice']);
      $sheet->setCellValue('B'.$writeRow, $item['stylecode']);
      $sheet->setCellValue('C'.$writeRow, $item['quantity']);
      $sheet->setCellValue('D'.$writeRow, $item['reference']);

      $writeColumn=5;

      foreach($uniqueComponents as $uniqueComponent){
        foreach($item['components'] as $component){
          if($component['name']==$uniqueComponent){
            $sheet->setCellValueByColumnAndRow($writeColumn, $writeRow, $component['weight']);
          }
        }

        $writeColumn++;
      }

      $writeRow++;
    }

    $writer = new Xlsx($spreadsheet);

    $writer->save($saveAs);

    return true;
  }
}
?>