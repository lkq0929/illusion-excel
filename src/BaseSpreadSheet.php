<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/8/28
 * Time: 14:01
 */

namespace illusion\excel;


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class BaseSpreadSheet
{
    public $phpOffice;
    
    public function __construct()
    {
        $this->phpOffice = new Spreadsheet();
    }
    
    /**
     * 写入数据
     *
     * @param array $values
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function write(array $values): Spreadsheet
    {
        $columnCount = isset($values[0]) ? \count($values[0]) : 0;
        $columns     = self::generateColumn($columnCount);
        foreach ($values as $row => $value) {
            foreach ($value as $column => $cellValue) {
                $this->phpOffice->setActiveSheetIndex(0)->setCellValue($columns[$column] . ($row + 1), $cellValue);
            }
        }
        
        return $this->phpOffice;
    }
    
    /**
     * 读取数据
     *
     * @param string $path
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function read(string $path): array
    {
        $tempFile = IOFactory::load($path);
        $values   = $tempFile->getActiveSheet()->toArray('', true, true, false);
        
        return $values;
    }
    
    /**
     * 写对对应后缀的电子表格文件
     *
     * @param string $type
     * @return IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writerToType(string $type): IWriter
    {
        return IOFactory::createWriter($this->phpOffice, ucfirst($type));
    }
    
    /**
     * 设置浏览器返回内容类型以及信息
     *
     * @param string $spreadSheetName
     * @param string $type
     */
    public function setHeader(string $spreadSheetName, string $type)
    {
        $mime = $this->getHeaderMime($type);
        \Yii::$app->response->setDownloadHeaders($spreadSheetName . '.' . strtolower($type), $mime)->send();
    }
    
    /**
     * 获取浏览器传输内容类型
     *
     * @param string $ext
     * @return string
     */
    public function getHeaderMime(string $ext): string
    {
        $mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        if (\in_array(ucfirst($ext), ['Xls', 'Csv'])) {
            $mime = 'application/vnd.ms-excel';
        }
        
        return $mime;
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
     * 将excel每一列与数据表中的属性对应
     *
     * @param array $rawData
     * @param array $attributes
     * @return array
     */
    public static function columnToAttribute(array $rawData, array $attributes): array
    {
        $transformData = [];
        $attribute     = [];
        if (!isset($rawData[0]) || (\count($rawData[0]) !== \count($attributes))) {
            throw new \InvalidArgumentException('INVALID PARAMS');
        }
        
        foreach ($rawData as $key => $values) {
            if ($key === 0) continue;
            foreach ($values as $column => $value) {
                $attribute[$attributes[$column]] = $value;
            }
            $transformData[] = $attribute;
        }
        
        return $transformData;
    }
    
    /**
     * excel原生数据转换成数据表里面的对应字段（顺序一致）
     *
     * @param array $rawValues
     * @param array $attributes
     * @return array
     */
    public function columnTo(array $rawValues = [], array $attributes = []): array
    {
        return static::columnToAttribute($rawValues, $attributes);
    }
    
    /**
     *  使用完成后释放内存
     */
    public function freeMemory()
    {
        $this->phpOffice->disconnectWorksheets();
        unset($this->phpOffice);
    }
    
}