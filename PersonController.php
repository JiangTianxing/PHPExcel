<?php

namespace App\Http\Controllers;

use App\Models\Person;

class PersonController extends Controller
{

    public function export()
    {
        $data = Person::getAllRooms();
        foreach ($data as $persons) {
            $this->createNewFile($persons);
        }
    }

    private function createNewFile($persons)
    {
        $template = storage_path('app/public/test.xls');
        $targetFile = storage_path("app/public/{$persons[0]['room_id']}.xls");

        //$objPHPExcel = new PHPExcel();                        //初始化PHPExcel(),不使用模板
        $objPHPExcel = \PHPExcel_IOFactory::load($template);     //加载excel文件,设置模板

        $objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);  //设置保存版本格式

        //接下来就是写数据到表格里面去
        $objActSheet = $objPHPExcel->getActiveSheet();
        $personNum   = count($persons);
        // 如果大于6条
        if ($personNum > 6) {
            for ($i = 0; $i < $personNum - 6; $i++) {
                $objActSheet->insertNewRowBefore(13);
            }
        }
        // 如果小于6条
        if ($personNum < 6) {
            $removeFrom = 12;
            for ($i = 0; $i < 6 - $personNum; $i++) {
                $objActSheet->removeRow($removeFrom);
            }
        }
        $personId = 2;
        $colIndex = 13;
        foreach ($persons as $person) {
            if ($person['relationship'] == '户主') {
                $this->display($objActSheet, $person, 12, 1);
            } else {
                $this->display($objActSheet, $person, $colIndex++, $personId++);
            }
        }

        $objWriter->save($targetFile);
    }

    private function display(\PHPExcel_Worksheet $objActSheet, $person, $colIndex, $personId)
    {
        $objActSheet->setCellValue('A' . $colIndex, $personId . " ");
        $objActSheet->setCellValue('B' . $colIndex, $person['name']);
        $objActSheet->setCellValue('C' . $colIndex, $person['gender']);
        $objActSheet->mergeCells('C' . $colIndex . ':D' . $colIndex);
        $objActSheet->setCellValue('E' . $colIndex, $person['id_num'] . " ");
        $objActSheet->mergeCells('E' . $colIndex . ':F' . $colIndex);
        $objActSheet->setCellValue('G' . $colIndex, $person['relationship']);
        $objActSheet->setCellValue('H' . $colIndex, $person['nationality']);
    }
}
