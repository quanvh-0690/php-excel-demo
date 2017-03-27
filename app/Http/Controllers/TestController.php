<?php

namespace App\Http\Controllers;

use App\Province;
use Illuminate\Http\Request;

class TestController extends Controller
{
    public function index()
    {
        $provinces = Province::with('districts')->get();
        $objPHPExcel = new \PHPExcel();
    
        foreach ($provinces as $key => $province) {
            $objPHPExcel->createSheet();
            $activeSheet = $objPHPExcel->setActiveSheetIndex($key);
            $activeSheet->setTitle($province->name);
            $activeSheet->setCellValue('A1', 'Quận/Huyện')
                ->setCellValue('B1', 'Xã/Phường')
                ->setCellValue('C1', 'Kinh độ, vĩ độ');
            $i = 2;
            $j = 2;
            foreach ($province->districts as $district) {
                $activeSheet->setCellValue("A$i", $district->type . ' ' . $district->name);
                foreach ($district->wards as $ward) {
                    $activeSheet->setCellValue("B$j", $ward->type . ' ' . $ward->name);
                    $activeSheet->setCellValue("C$j", $ward->location);
                    $j++;
                }
                $rowMerge = $j - 1;
                if ($i < $rowMerge) {
                    $activeSheet->mergeCells("A$i:A$rowMerge");
                }
                $i = $j;
            }
        }
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save(base_path('result.xlsx'));
    }
}
