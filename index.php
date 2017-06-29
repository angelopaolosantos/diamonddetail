<?php 
require __DIR__ . '/vendor/autoload.php';
require '/classes/DiamondDetail.php';

$success=null;
$errors=array();

if($_SERVER['REQUEST_METHOD'] == "POST"){
  if($_FILES['file']['tmp_name']==''){
    $errors[]="No excel file was uploaded.";
  }
  if($_POST['start']=='' && !is_int($_POST['start'])){
    $errors[]="Starting row required and must be an integer.";
  }
  if($_POST['end']=='' && !is_int($_POST['end'])){
    $errors[]="Ending row required and must be an integer.";
  }
  if($_POST['component-start']=='' && !is_int($_POST['component-start'])){
    $errors[]="Component Starting column required and must be an integer.";
  }
  if($_POST['component-end']=='' && !is_int($_POST['component-end'])){
    $errors[]="Component Ending column required and must be an integer.";
  }

  if (count($errors)==0){
    $diamondDetail = new DiamondDetail($_POST, $_FILES);
    $data=$diamondDetail->readExcelFile();
  
    if($diamondDetail->writeExcelFile($data)){
      $success="Success! New excel file was created.";
    }
  }
}


?>
<!DOCTYPE html>
<html>
<head>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://unpkg.com/purecss@0.6.2/build/pure-min.css" integrity="sha384-UQiGfs9ICog+LwheBSRCt1o5cbyKIHbwjWscjemyBMT9YCUMZffs6UqUTd0hObXD" crossorigin="anonymous">
  <link rel="stylesheet" type="text/css" href="style.css">
  <title>Billion Sail - Diamond Detail</title>
</head>
<body>
<h1>Billion Sail - Diamond Detail</h1>
  <?php
    if($success){
  ?>
    <div class="isa_success">
     <i class="fa fa-check"></i><?=$success?></div>
  <?php
    }

    if($errors){
  ?>
    <div class="isa_error">
    <ul>
  <?php
    foreach($errors as $error){
      echo "<li><i class='fa fa-exclamation'></i>".$error."</li>";
    }
  ?>
    </ul>
     </div>
    <?php
    }
  ?>
  <form class="pure-form pure-form-aligned" action="<?php echo $_SERVER['PHP_SELF']; ?>" method="post" enctype="multipart/form-data">
    <fieldset>

      <div class="pure-control-group">
        <label for="file">Excel File</label>
        <input id="file" name="file" type="file" placeholder="excel.xls">
        <span class="pure-form-message-inline">This is a required field.</span>
      </div>

      <div class="pure-control-group">
        <label for="start">Starting Row</label>
        <input id="start" name="start" type="text" placeholder="1">
        <span class="pure-form-message-inline">This is a required field.</span>
      </div>

      <div class="pure-control-group">
        <label for="end">Ending Row</label>
        <input id="end" name="end" type="text" placeholder="100">
        <span class="pure-form-message-inline">This is a required field.</span>
      </div>

      <div class="pure-control-group">
        <label for="component-start">Starting Component Column</label>
        <input id="component-start" name="component-start" type="text" placeholder="5">
        <span class="pure-form-message-inline">This is a required field.</span>
      </div>

      <div class="pure-control-group">
        <label for="component-end">Ending Component Column</label>
        <input id="component-end" name="component-end" type="text" placeholder="100">
        <span class="pure-form-message-inline">This is a required field.</span>
      </div>

      <div class="pure-controls">
        <button type="submit" class="pure-button pure-button-primary">Submit</button>
      </div>

    </fieldset>
  </form>
</body>
</html>