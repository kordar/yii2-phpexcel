<?php

namespace kordar\phpexcel;

use kordar\phpexcel\widget\OutputCsv;
use kordar\phpexcel\widget\OutputOds;

/**
 * This is just an example.
 */
class AutoloadExample extends \yii\base\Widget
{
    /**
     * @var \yii\db\Query $query
     */
    public $query = null;

    public function run()
    {
        try {





            // $objPHPExcel = new \PHPExcel();
            /**
             * @var \PHPExcel_Writer_CSV $csvObj
             */
            /*$csvObj = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
            $csvObj->setDelimiter(',')
                   ->setEnclosure('"')
                   ->setLineEnding("\r\n")
                   ->setSheetIndex(0);
            $csvObj->save('456.csv');*/

            $title = ['ID', 'Business', 'Title', 'adID', 'phone', 'investMoney', 'reg', 'settlement', 'expend', 'invest', 'complex', 'market', 'bc', 'dsaf'];
            OutputOds::widget(['filename' => '123456', 'query' => $this->query, 'sheetTitle' => $title]);

            // die;
            set_time_limit(600);
            $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_apc;
            $cacheSettings = array( 'cacheTime' => 600 );
            \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

            // Create new PHPExcel object
            $objPHPExcel = new \PHPExcel();

            // Set document properties
            $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
                ->setLastModifiedBy("Maarten Balliauw")
                ->setTitle("Office 2007 XLSX Test Document")
                ->setSubject("Office 2007 XLSX Test Document")
                ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                ->setKeywords("office 2007 openxml php")
                ->setCategory("Test result file");

            $obj = $objPHPExcel->setActiveSheetIndex(0);

            foreach ($this->query->each() as $key => $val) {
                // Add some data
                $num = $key + 1;
                $obj->setCellValue('A' . $num, $val['id'])
                    ->setCellValue('B' . $num, $val['business'])
                    ->setCellValue('C' . $num, $val['title'])
                    ->setCellValue('D' . $num, $val['adID'])
                    ->setCellValue('E' . $num, $val['phone'])
                    ->setCellValue('F' . $num, $val['investTime'])
                    ->setCellValue('G' . $num, $val['investMoney'])
                    ->setCellValue('H' . $num, $val['regMoney'])
                    ->setCellValue('I' . $num, $val['settlementMoney'])
                    ->setCellValue('J' . $num, $val['expenditureMoney'])
                    ->setCellValue('K' . $num, $val['investTerm'])
                    ->setCellValue('L' . $num, $val['isComplex'])
                    ->setCellValue('M' . $num, $val['market'])
                    ->setCellValue('N' . $num, $val['bc']);
            }


            // Rename worksheet
            $objPHPExcel->getActiveSheet()->setTitle('Simple');


            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $objPHPExcel->setActiveSheetIndex(0);


            // Redirect output to a client’s web browser (OpenDocument)
            header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');
            header('Content-Disposition: attachment;filename="01simple.ods"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
            header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header ('Pragma: public'); // HTTP/1.0

            $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'OpenDocument');
            $objWriter->save('php://output');
            exit;

        } catch (\Exception $e) {

        }




    }


    public function csv()
    {
        $file_name="CSV".date("mdHis",time()).".csv";
        header ( 'Content-Type: application/vnd.ms-excel' );
        header ( 'Content-Disposition: attachment;filename='.$file_name );
        header ( 'Cache-Control: max-age=0' );
        $file = fopen('php://output',"a");
        $limit=1000;
        $calc=0;

        foreach ($this->query->each() as $key => $val) {
            // Add some data
            $calc++;
            if($limit==$calc){
                ob_flush();
                flush();
                $calc=0;
            }
            foreach ($val as $t){
                $tarr[]=iconv('UTF-8', 'GB2312//IGNORE',$t);
            }
            fputcsv($file,$tarr);
            unset($tarr);



        }

        fclose($file);

        exit();
    }


}
