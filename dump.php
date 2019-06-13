<?php

define('DB',[
    'host' =>  "127.0.0.1",
    'port' => 3306,
    'login' =>  'root' ,
    'password' => '1q2w3e4r$1' ,
    'db' => 'db2',
]);

define('ROOT', dirname(__FILE__));
define('ROOT_LIBRARY', ROOT);
define('OUTPUT_FILE', ROOT.'/OUT.xlsx');

include ROOT_LIBRARY . '/PHPExcel.php';
include ROOT_LIBRARY . '/PHPExcel/Writer/Excel2007.php';
include ROOT_LIBRARY . '/PHPExcel/Writer/Excel5.php';
include ROOT_LIBRARY . '/PHPExcel/IOFactory.php';

define('EXCEL_NAMES',['A',
      'B',
      'C',
      'D',
      'E',
      'F',
      'G',
      'H',
      'I',
      'J',
      'K',
      'L',
      'M',
      'N',
      'O',
      'P',
      'Q',
      'R',
      'S',
      'T',
      'U',
      'V',
      'W',
      'X',
      'Y',
      'Z',
      'AA',
      'AB',
      'AC',
      'AD',
      'AE',
      'AF',
      'AG',
      'AH',
      'AI',
      'AJ',
      'AK',
      'AL',
      'AM',
      'AN',
      'AO',
      'AP',
      'AQ',
      'AR',
      'AS',
      'AT',
      'AU',
      'AV',
      'AW',
      'AX',
      'AY',
      'AZ'
]);



$db = DB;

//如果需要指定几张表，则枚举。不需要指定则留空。
$table_names = [];
// $table_names[] = 'biz_ticket'; 

function db_query($conn, $sql) {
    $result = mysqli_query($conn,   $sql);
    $ret = [];
    if ($result->num_rows > 0) {
        while($row = $result->fetch_assoc()) {
            $ret[] = $row;
        }
    }
    return $ret;
}

function get_table_info($conn, $value){
    if (isset($value['TABLE_NAME'])) {
        $table_item = [];
        $table_name =  $value['TABLE_NAME'];
        $db = DB;
        $sql = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='" . $db['db'] . "' AND  TABLE_NAME='".$table_name."' ;";
        $res = db_query($conn,  $sql);
        $value['COLUMNS'] = $res;
        return $value;
    }
    return [];
}

function exportExcel($table_list){
    $HEAD_COLOR = '8EAADC';
    $objExcel = new \PHPExcel();
    $objExcel->getProperties()->setCreator("andy");
    $objExcel->getProperties()->setLastModifiedBy("andy");
    $objExcel->getProperties()->setTitle("Office 2003 XLS Test Document");
    $objExcel->getProperties()->setSubject("Office 2003 XLS Test Document");
    $objExcel->getProperties()
      ->setDescription("Test document for Office 2003 XLS, generated using PHP classes.");
    $objExcel->getProperties()->setKeywords("office 2003 openxml php");
    $objExcel->getProperties()->setCategory("Test result file");
    $objExcel->setActiveSheetIndex(0);

    $cnames = ['表名','字段名','字段描述','类型','长度','小数','索引','外键','非空','缺省','说明'];
    $widthArr = [30,20,20,10,10,10,10,10,10];

    $rowIndex = 1;// 第一行
    foreach ($table_list as $key => $value) {
        $TABLE_NAME = $value['TABLE_NAME'];
        $TABLE_COMMENT =  $value['TABLE_COMMENT'];
        if (count($cnames) > 0) {
            $row_start = EXCEL_NAMES[0] . $rowIndex;
            $row_end = EXCEL_NAMES[count($cnames) - 1] . $rowIndex;
            foreach ($cnames as $ck => $cv) {//渲染列名
                $objExcel->getActiveSheet()->setCellValue(EXCEL_NAMES[$ck] . $rowIndex, $cv);
               
            }
            $objExcel->getActiveSheet()->getStyle( $row_start.':'.$row_end)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
            $objExcel->getActiveSheet()->getStyle( $row_start.':'.$row_end)->getFill()->getStartColor()->setARGB('FF'.$HEAD_COLOR);
        }
        $rowIndex = $rowIndex + 1;
        $COLUMNS = $value['COLUMNS'];
        $cvalues = [];
        foreach ($COLUMNS as $k => $v) {
            $COLUMN_NAME = $v['COLUMN_NAME'];
            $COLUMN_DEFAULT = $v['COLUMN_DEFAULT'];
            $IS_NULLABLE = $v['IS_NULLABLE'];
            $DATA_TYPE = $v['DATA_TYPE'];
            $CHARACTER_MAXIMUM_LENGTH = $v['CHARACTER_MAXIMUM_LENGTH'];
            $CHARACTER_OCTET_LENGTH = $v['CHARACTER_OCTET_LENGTH'];
            $NUMERIC_PRECISION = $v['NUMERIC_PRECISION'];
            $NUMERIC_SCALE = $v['NUMERIC_SCALE'];
            $DATETIME_PRECISION = $v['DATETIME_PRECISION'];
            $COLUMN_TYPE = $v['COLUMN_TYPE'];
            $COLUMN_KEY = $v['COLUMN_KEY'];
            $COLUMN_COMMENT = $v['COLUMN_COMMENT'];
            $LEN = $NUMERIC_PRECISION;
            if ($DATA_TYPE == 'varchar') {
                $LEN = $CHARACTER_MAXIMUM_LENGTH;
            }
            $INDEX = '';
            if($COLUMN_KEY == 'PRI')  $INDEX = 'P';
            if($COLUMN_KEY == 'MUL')  $INDEX = 'U';
            $NULLABLE = '';
            if($IS_NULLABLE == 'YES') $NULLABLE = 'Y';
            if($IS_NULLABLE == 'NO') $NULLABLE = 'N';
            $cvalues[] = [$TABLE_NAME."\n".$TABLE_COMMENT,$COLUMN_NAME,$COLUMN_COMMENT, $DATA_TYPE, $LEN,$NUMERIC_SCALE,$INDEX,'',$NULLABLE,$COLUMN_DEFAULT,'' ];
        }

        if (count($cvalues) > 0) {
            $merge_start = EXCEL_NAMES[0] . $rowIndex;
            $merge_end = EXCEL_NAMES[0] . ($rowIndex + count($cvalues) - 1);
            $center_start = 'E'.$rowIndex;
            $center_end = 'J'.($rowIndex + count($cvalues) - 1);
            foreach ($cvalues as $k => $v) {
                if (empty($merge_start)) {
                    $merge_start = EXCEL_NAMES[0] . $rowIndex;
                }
                foreach ($v as $ck => $cv) {//渲染列名
               
                    $objExcel->getActiveSheet()->setCellValue(EXCEL_NAMES[$ck] . $rowIndex, $cv);
                }
                $rowIndex = $rowIndex + 1;
            }
            $objExcel->getActiveSheet()->mergeCells($merge_start.':'.$merge_end);//->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);  
            $objExcel->getActiveSheet()->getStyle($merge_start.':'.$merge_end)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $objExcel->getActiveSheet()->getStyle($merge_start.':'.$merge_end)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $objExcel->getActiveSheet()->getStyle($merge_start.':'.$merge_end)->getAlignment()->setWrapText(true);
            $objExcel->getActiveSheet()->getStyle($center_start.':'.$center_end)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $objExcel->getActiveSheet()->getStyle($center_start.':'.$center_end)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        }
        $rowIndex = $rowIndex + 1;
    }
   
        // 高置列的宽度
    if (count($widthArr) > 0) {//默认50
      $len = count($widthArr);
      for ($i = 0; $i<$len; $i++){
        $cellKey = EXCEL_NAMES[$i];
        $w = $widthArr[$i];
        $objExcel->getActiveSheet()->getColumnDimension($cellKey)->setWidth($w);
      }
    }
     $objExcel->getActiveSheet()
      ->getHeaderFooter()
      ->setOddHeader('&L&BPersonal cash register&RPrinted on &D');
    $objExcel->getActiveSheet()
      ->getHeaderFooter()
      ->setOddFooter('&L&B' . $objExcel->getProperties()
          ->getTitle() . '&RPage &P of &N');
    // 设置页方向和规模
    $objExcel->getActiveSheet()
      ->getPageSetup()
      ->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
    $objExcel->getActiveSheet()->getPageSetup()->setPaperSize(\PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
    $objExcel->setActiveSheetIndex(0);
    $timestamp = time();

    $objWriter = \PHPExcel_IOFactory::createWriter($objExcel, 'Excel2007');
    $objWriter->save(OUTPUT_FILE);
    exit;
}

$use_mysqli = function_exists("mysqli_connect");
if ($use_mysqli) {
    $conn = mysqli_connect($db['host'], $db["login"], $db["password"]);
    mysqli_set_charset($conn, 'utf8');
    $errno_c = mysqli_connect_errno($conn);
    if($errno_c > 0) {
           echo "连接失败";
        exit;
    }
    if(($errno_c <= 0) && ( $db["db"] != "" )) {
        $res = mysqli_select_db($conn, $db["db"] );
        $errno_c = mysqli_errno($conn);
    }
    $sql = "SELECT TABLE_NAME,TABLE_ROWS,TABLE_COMMENT FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='".$db['db']."';";
    $res = db_query($conn,   $sql);
    $table_list = [];
    foreach ($res as $key => $value) {
        $tname = $value['TABLE_NAME'];
        if (empty($table_names)) {
              $table_list[] = get_table_info($conn,$value);
        }elseif (in_array($tname, $table_names)) {
           $table_list[] = get_table_info($conn,$value);
        }
    }
    exportExcel($table_list);
}

