<?php
/**
 * Created by PhpStorm.
 *
 * Date: 2018/8/7
 * Time: 14:23
 */
namespace illusion\excel;

use illusion\excel\services\ExportService;
use illusion\excel\services\ImportService;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use yii\base\Component;

class Excel extends Component {
    
    public $tempPath; //文件的暂存路径
    
    public function __construct($tempPath = '')
    {
        parent::__construct([]);
        $this->tempPath = $tempPath;
    }
    
    /**
     * 填充数据到excel
     *
     * auth: lkqlink@163.com
     *
     * @param array $cellData
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function export(array $cellData = []): Spreadsheet
    {
        $sheetInstance = new Spreadsheet();
        $sheetInstance->setActiveSheetIndex(0);
        $defaultWorkSheet = $sheetInstance->getActiveSheet();
        $columns          = ExportService::generateColumn(\count($cellData[0]));
        foreach ($cellData as $row => $cellValues) {
            foreach ($cellValues as $column => $cellValue) {
                $defaultWorkSheet->getColumnDimension($columns[$column])->setWidth('15');
                $defaultWorkSheet->setCellValue($columns[$column] . ($row + 1), $cellValue);
            }
        }
        
        return $sheetInstance;
    }
    
    /**
     * 直接下载excel文件
     *
     * auth: lkqlink@163.com
     *
     * @param $spreadSheet
     * @param string $fileName
     * @param string $ext
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function download($spreadSheet, string $fileName, string $ext)
    {
        $exportFile = IOFactory::createWriter($spreadSheet, $ext);
        $extToMime  = ExportService::getExtToMime($ext);
        header('Content-Type:' . $extToMime);
        header('Content-Disposition:attachment;filename="' . $fileName . '.' . $ext . '"');
        $exportFile->save('php://output');
    }
    
    /**
     * 数据导入
     *
     * auth: lkqlink@163.com
     *
     * @param array $attributes
     * @param string $isRaw
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function import(array $attributes = [], string $isRaw = 'raw'): array
    {
        $tempFile = IOFactory::load($this->tempPath);
        $rawData  = $tempFile->getActiveSheet()->toArray('', true, true, false);
        $results = $isRaw === 'raw' ? $rawData : ImportService::columnToAttribute($rawData, $attributes);
        
        return $results;
    }
}