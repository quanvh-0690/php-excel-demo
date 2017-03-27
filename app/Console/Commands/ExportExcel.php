<?php

namespace App\Console\Commands;

use App\Province;
use Illuminate\Console\Command;

class ExportExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'db:export_to_excel';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Export address to excel';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function handle()
    {
        $provinces = Province::with('districts')->get();
        $objPHPExcel = new \PHPExcel();
    
        foreach ($provinces as $key => $province) {
            $objPHPExcel->createSheet(); // tạo 1 sheet mới
            $activeSheet = $objPHPExcel->setActiveSheetIndex($key);
            $activeSheet->setTitle($province->name); // đặt tên sheet là tên tỉnh
            $activeSheet->setCellValue('A1', 'Quận/Huyện')
                ->setCellValue('B1', 'Xã/Phường')
                ->setCellValue('C1', 'Kinh độ, vĩ độ'); // set title cho dòng đầu tiên
            $i = 2;
            $j = 2;
            foreach ($province->districts as $district) {
                $activeSheet->setCellValue("A$i", $district->type . ' ' . $district->name); // set tên quận/huyện
                foreach ($district->wards as $ward) {
                    $activeSheet->setCellValue("B$j", $ward->type . ' ' . $ward->name); // tương ứng mỗi quận huyện set tên xã/phường
                    $activeSheet->setCellValue("C$j", $ward->location);
                    $j++;
                }
                $rowMerge = $j - 1;
                if ($i < $rowMerge) {
                    $activeSheet->mergeCells("A$i:A$rowMerge"); // merge các cell có cùng 1 quận/huyện
                }
                $i = $j;
            }
            
            foreach (range('A', 'C') as $columnId) {
                $activeSheet->getColumnDimension($columnId)->setAutoSize(true);
            }
        }
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save(base_path('result.xlsx'));
    }
}
