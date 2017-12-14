<?php
namespace kordar\phpexcel\widget;

class OutputOds extends QueryIterator
{
    protected $fileType = 'ODS';

    public $row = 1;

    public function run()
    {
        /**
         * 创建单页 sheet 应用
         */
        if (is_string($this->query)) {
            $this->createSheet($this->query, 0, $this->sheetTitle);
        }

    }

    protected function setSheetTitles($titles = [])
    {
        if (!empty($titles)) {
            foreach ($titles as $col => $title) {
                $cell = $this->stringFromColumnIndex($col, $this->row);
                $this->activeSheet->setCellValue($cell, $title);
            }
            $this->row ++;
        }
    }

    public function createSheet(\yii\db\Query $query, $index = 0, $titles = [])
    {
        try {
            $this->activeSheet = $this->handlePHPExcel->setActiveSheetIndex($index);
            $this->setSheetTitles($titles);

            foreach ($query->each() as $each) {
                $values = array_values($each);
                foreach ($values as $key => $val) {
                    $cell = $this->stringFromColumnIndex($key, $this->row);
                    $this->activeSheet->setCellValue($cell, $val);
                }
                $this->row ++;
            }
        } catch (\Exception $e) {
            echo $e->getMessage() . PHP_EOL;
        }
    }

}