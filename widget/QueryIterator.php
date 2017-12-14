<?php
namespace kordar\phpexcel\widget;

use kordar\phpexcel\helper\Header;
use yii\base\Widget;

abstract class QueryIterator extends Widget implements FileTypeInterface
{
    /**
     * @var \yii\db\Query $query
     */
    public $query;

    public $filename = '';
    public $sheetTitle = [];

    protected $fileType = 'Excel5';

    protected $document = [
        FileTypeInterface::FILE_TYPE_ODS => 'OpenDocument',
    ];

    /**
     * @var \PHPExcel $handlePHPExcel
     */
    protected $handlePHPExcel = null;

    protected $activeSheet = null;

    protected $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_apc;
    protected $cacheSettings = ['cacheTime' => 600];

    public function init()
    {
        if ($this->fileType === FileTypeInterface::FILE_TYPE_CSV) {
            Header::csv($this->filename);
            return true;
        }

        $this->handlePHPExcel = new \PHPExcel();

        /**
         * set document properties
         * like "creator" "lastModitifiedBy" "title" "subject" ....
         * default equal to ""
         */
        $this->setDocumentProperties();

        $this->on(self::EVENT_AFTER_RUN, function () {

            switch ($this->fileType) {
                case FileTypeInterface::FILE_TYPE_ODS:
                    Header::ods($this->filename);
                    break;
            }



        });

    }

    protected function output()
    {
        try {
            $objWriter = \PHPExcel_IOFactory::createWriter($this->handlePHPExcel, $this->document[$this->fileType]);
            $objWriter->save('php://output');
        } catch (\Exception $e) {
            echo 'Export Failed' . PHP_EOL;
        }
    }

    abstract public function createSheet(\yii\db\Query $query, $index = 0, $titles = []);

    /**
     * @param $col
     * @param $row
     * @return string
     *
     */
    public function stringFromColumnIndex($col, $row)
    {
        return \PHPExcel_Cell::stringFromColumnIndex($col) . $row;
    }


    // 文档属性
    public $properties = [];
    // 设置文档属性
    protected function setDocumentProperties()
    {
        $properties = array_merge(['createor'=>'','lastModifiedBy'=>'','title'=>'','subject'=>'','description'=>'','keywords'=>'','category'=>''], $this->properties);
        $this->handlePHPExcel->getProperties()->setCreator($properties['createor'])
            ->setLastModifiedBy($properties['lastModifiedBy'])
            ->setTitle($properties['title'])
            ->setSubject($properties['subject'])
            ->setDescription($properties['description'])
            ->setKeywords($properties['keywords'])
            ->setCategory($properties['category']);
    }

}