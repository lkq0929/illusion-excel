<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/8/31
 * Time: 11:04
 */

namespace illusion\excel;


use PhpOffice\PhpSpreadsheet\Spreadsheet;

class WorkSheet
{
    public $spreadsheet;
    
    public function __construct(Spreadsheet $spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
    }
    
    /**
     * 清除第一张工作表
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function clearFirstSheet()
    {
        if (isset($this->spreadsheet->getAllSheets()[0])) {
            $this->spreadsheet->removeSheetByIndex(0);
        }
    }
    
    /**
     * 获取当前的活动工作表
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getCurrentActiveSheet(): \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    {
        return $this->spreadsheet->getActiveSheet();
    }
    
    /**
     * 导出多个电子工作表的电子表格
     *
     * @param array $mulValues
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function mulWorksheet(array $mulValues): Spreadsheet
    {
        $pIndex = -1;
        $this->clearFirstSheet();
        foreach ($mulValues as $sheetName => $mulValue) {
            $this->writeSingleSheet($mulValue, $sheetName, ++$pIndex);
        }
        return $this->spreadsheet;
    }
    
    /**
     * 给当前活动的工作表填充数据
     *
     * @param array $values
     * @param string $sheetName
     * @param int $pIndex
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function writeSingleSheet(array $values, string $sheetName = 'worksheet', $pIndex = 0): Spreadsheet
    {
        $spreadsheet = $this->spreadsheet;
        if (!isset($values[0])) {
            throw new \InvalidArgumentException('INVALID PARAMS');
        }
        $columns = self::generateColumn(\count($values[0]));
        $spreadsheet->createSheet();
        $spreadsheet->setActiveSheetIndex($pIndex);
        foreach ($values as $row => $value) {
            foreach ($value as $column => $cellValue) {
                $this->getCurrentActiveSheet()->setCellValue($columns[$column] . ($row + 1), $cellValue);
            }
        }
        $this->getCurrentActiveSheet()->setTitle($sheetName);
        
        return $spreadsheet;
    }
    
    /**
     * 根据一行的个数生成excel对应的title元素个数
     *
     * EXAMPLE: $number = 27,  可以生成 ['A'....'Z', 'AA']
     *
     * @param int $number
     * @return array
     */
    public static function generateColumn(int $number = 0): array
    {
        $baseColumn = range('A', 'Z');
        $columns    = $baseColumn;
        if ($number > 26) {
            $subColumn = $number - 26;
            $mulColumn = floor($subColumn / 26);  //向下取整
            $surColumn = $number - (($mulColumn + 1) * 26);
            if ($mulColumn <= 0) {
                for ($column = 0; $column < $subColumn; $column++) {
                    $columns[] = $baseColumn[$mulColumn] . $baseColumn[$column];
                }
            } else {
                for ($layer = 0; $layer < $mulColumn; $layer++) {
                    for ($column = 0; $column < 26; $column++) {
                        $columns[] = $baseColumn[$layer] . $baseColumn[$column];
                    }
                    if ($layer + 2 > $mulColumn) { //最后一层求余
                        for ($column = 0; $column < $surColumn; $column++) {
                            $columns[] = $baseColumn[$layer + 1] . $baseColumn[$column];
                        }
                    }
                }
            }
        }
        
        return $columns;
    }
    
    /**
     * 读取单个电子表数据
     *
     * @param Spreadsheet $spreadsheet
     * @param int $pIndex
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function readSingleSheet(Spreadsheet $spreadsheet, int $pIndex = 0): array
    {
        $spreadsheet->setActiveSheetIndex($pIndex);
        $rawData = $spreadsheet->getActiveSheet()->toArray('', true, true, false);
        
        return $rawData;
    }
    
    /**
     * 读取多个电子表数据
     *
     * @param Spreadsheet $spreadsheet
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function mulSheetData(Spreadsheet $spreadsheet): array
    {
        $values     = [];
        $sheetNames = $spreadsheet->getSheetNames();
        foreach ($sheetNames as $pIndex => $sheetName) {
            $values[$sheetName] = $this->readSingleSheet($spreadsheet, $pIndex);
        }
        
        if (\count($values) === 1) {  //单sheet工作簿数据降维处理
            $values = array_shift($values);
        }
        
        return $values;
    }
    
}