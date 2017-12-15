<?php
namespace kordar\phpexcel\web;

use kordar\phpexcel\helper\Header;

class Document implements DocumentInterface
{
    public $type = 'excel5';

    public $name = '';

    /**
     * @var \PHPExcel $handlePHPExcel
     */
    protected $handlePHPExcel = null;

    protected $types = [
        DocumentInterface::DOCUMENT_TYPE_ODS => 'OpenDocument',
    ];

    public function output()
    {
        try {
            $objWriter = \PHPExcel_IOFactory::createWriter($this->handlePHPExcel, $this->types[$this->type]);
            $objWriter->save('php://output');
        } catch (\Exception $e) {
            echo 'Export Failed' . PHP_EOL;
        }
    }

}