<?php
namespace kordar\phpexcel\helper;

class Header
{

    public static function csv($filename)
    {
        header ( 'Content-Type: application/vnd.ms-excel' );
        header ( 'Content-Disposition: attachment;filename="' . $filename . '.csv"');
        header ( 'Cache-Control: max-age=0' );
    }

    public static function ods($filename)
    {
        // Redirect output to a client’s web browser (OpenDocument)
        header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');
        header('Content-Disposition: attachment;filename="' . $filename . '.ods"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
    }


    public static function excel($filename, $type = 'Excel5')
    {
        if( $type == "Excel5" ) {
            // Redirect output to a client’s web browser (Excel5)
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . $filename . '.xls"');
        } else {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
        }

        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
    }

    public static function pdf($filename)
    {
        // Redirect output to a client’s web browser (PDF)
        header('Content-Type: application/pdf');
        header('Content-Disposition: attachment;filename="' . $filename . '.pdf"');
        header('Cache-Control: max-age=0');
    }

}