<?php
/**
 * Created by PhpStorm.
 *
 * auth: lkqlink@163.com
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
    
    
    public function __construct(array $config = [])
    {
        parent::__construct($config);
    }
    
    /**
     * 填充数据到excel
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
        $columnCount      = isset($cellData[0]) ? \count($cellData[0]) : 0;
        $columns          = ExportService::generateColumn($columnCount);
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
     * @param $spreadSheet
     * @param string $fullFileName
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function download($spreadSheet, string $fullFileName)
    {
        $excel = ExportService::handleFileName($fullFileName);
        $exportFile = ExportService::writeData($spreadSheet, ucfirst($excel['ext']));
        $extToMime  = ExportService::getExtToMime($excel['ext']);
        header('Content-Type:' . $extToMime);
        header('Content-Disposition:attachment;filename="' . $fullFileName . '"');
        $exportFile->save('php://output');
    }
    
    /**
     * 生成的文件保存到本地
     *
     * @param $spreadSheet
     * @param string $attachPath  YII框架，默认的下载目录在we根目录下
     * @param string $fullFileName
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function store($spreadSheet, string $fullFileName, string $attachPath = './')
    {
        $excel      = ExportService::handleFileName($fullFileName);
        $exportFile = ExportService::writeData($spreadSheet, ucfirst($excel['ext']));
        $exportFile->save($attachPath . $fullFileName);
    }
    
    /**
     * 数据导入
     *
     * @param string $tempPath
     * @param array $attributes
     * @param string $isRaw
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function import(string $tempPath, array $attributes = [], string $isRaw = ''): array
    {
        $tempFile = IOFactory::load($tempPath);
        $rawData  = $tempFile->getActiveSheet()->toArray('', true, true, false);
        $results = $isRaw === 'raw' ? $rawData : ImportService::columnToAttribute($rawData, $attributes);
        
        return $results;
    }
}