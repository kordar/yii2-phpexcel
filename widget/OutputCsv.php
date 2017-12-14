<?php
namespace kordar\phpexcel\widget;


class OutputCsv extends QueryIterator
{
    protected $fileType = 'CSV';

    public function run()
    {
        $file = fopen('php://output',"a");
        fwrite($file, "\xEF\xBB\xBF");

        $limit=1000;
        $calc=0;

       if (!empty($this->sheetTitle)) {
            fputcsv($file, $this->sheetTitle);
            unset($tarr);
            $calc++;
        }

       foreach ($this->query->each() as $key => $val) {
            // Add some data
            $calc++;
            if($limit==$calc){
                ob_flush();
                flush();
                $calc=0;
            }
            foreach ($val as $t){
                $tarr[] = $t;
            }
            fputcsv($file, $tarr);
            unset($tarr);
        }
        fclose($file);
        exit();
    }
}