<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/9/26
 * Time: 17:06
 */

namespace illusion\excel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Worksheet
{
    /**
     * @var 工作表名称
     */
    private $sheetName;
    
    
    private $_spreadsheet;
    
    public function __construct(Spreadsheet $spreadsheet)
    {
        $this->_spreadsheet = $spreadsheet;
    }
    
    /**
     * 设置工作表的名称
     *
     * @param $sheetName
     */
    public function setSheetName($sheetName)
    {
        $this->sheetName = $sheetName;
    }
    
    /**
     * 根据列数计算需要占用到的列的字母表示
     *
     * @param $columnNum
     * @return array
     */
    public static function calculateColumnAlphabet($columnNum): array
    {
        $baseColumn = range('A', 'Z');
        $columns    = $baseColumn;
        if ($columnNum > 26) {
            $subColumn = $columnNum - 26;
            $mulColumn = floor($subColumn / 26);  //向下取整
            $surColumn = $columnNum - (($mulColumn + 1) * 26);
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
     * 删掉第一个工作表
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function clearFirstSheet()
    {
        if (isset($this->_spreadsheet->getAllSheets()[0])) {
            $this->_spreadsheet->removeSheetByIndex(0);
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
        return $this->_spreadsheet->getActiveSheet();
    }
    
    /**
     * 填充数据到工作表
     *
     * @param array $values
     * @param string $sheetName
     * @param int $pIndex
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function fillValuesToSheet(array $values, string $sheetName = 'worksheet', $pIndex = 0): Spreadsheet
    {
        $spreadsheet = $this->_spreadsheet;
        if (!isset($values[0])) {
            throw new \InvalidArgumentException('INVALID PARAMS');
        }
        $columns = self::calculateColumnAlphabet(\count($values[0]));
        $spreadsheet->createSheet();
        $spreadsheet->setActiveSheetIndex($pIndex);
        foreach ($values as $row => $value) {
            foreach ($value as $column => $cellValue) {
                $this->getCurrentActiveSheet()->getColumnDimension($columns[$column])->setAutoSize(true);
                $this->getCurrentActiveSheet()->setCellValue($columns[$column] . ($row + 1), $cellValue);
            }
        }
        $this->getCurrentActiveSheet()->setTitle($sheetName);
    
        return $spreadsheet;
    }
    
    /**
     * 填充数据到多个工作表
     *
     * @param array $values
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function fillValuesToSheets(array $values): Spreadsheet
    {
        $pIndex = -1;
        $this->clearFirstSheet();
        foreach ($values as $sheetName => $value) {
            $this->fillValuesToSheet($value, $sheetName, ++$pIndex);
        }
    
        return $this->_spreadsheet;
    }
    
    /**
     * 获取当前工作表数据
     *
     * @param Spreadsheet $spreadsheet
     * @param int $pIndex
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getSheetValues(Spreadsheet $spreadsheet, int $pIndex = 0): array
    {
        $spreadsheet->setActiveSheetIndex($pIndex);
        $rawData = $spreadsheet->getActiveSheet()->toArray('', true, true, false);
        
        return $rawData;
    }
    
    /**
     * 获取电子表中的所有工作表的数据
     *
     * @param Spreadsheet $spreadsheet
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getSheetsValues(Spreadsheet $spreadsheet): array
    {
        $values     = [];
        $sheetNames = $spreadsheet->getSheetNames();
        foreach ($sheetNames as $pIndex => $sheetName) {
            $values[$sheetName] = $this->getSheetValues($spreadsheet, $pIndex);
        }
        
        if (\count($values) === 1) {  //单sheet工作簿数据降维处理
            $values = array_shift($values);
        }
        
        return $values;
    }
}