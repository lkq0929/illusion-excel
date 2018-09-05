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
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class BaseSpreadSheet
{
    protected $phpOffice;
    private $worksheet;
    
    public function __construct()
    {
        $this->phpOffice = new Spreadsheet();
        $this->worksheet = new \illusion\excel\WorkSheet($this->phpOffice);
    }
    
    /**
     * 导出数据（默认单个sheet,三维数组对应多个工作表）
     *
     * @param array $values
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function write(array $values): Spreadsheet
    {
        
        $valuesDepth = self::arrayDepth($values);
        if ($valuesDepth <= 1 || $valuesDepth > 3) {
            throw new \InvalidArgumentException('INVALID　PARAMS');
        }
        
        if ($valuesDepth === 2) {
            $values = [$values];
        }
        
        return $this->worksheet->mulWorksheet($values);
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
        $spreadSheet = IOFactory::load($path); //就是电子表格实例
        $values      = $this->worksheet->mulSheetData($spreadSheet);
        
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
     * 将excel每一列与数据表中的属性对应
     *
     * @param array $rawData
     * @param array $attributes
     * @param int $filterByKey 根据第n列的数据过滤掉每行的空字符串
     * @return array
     */
    public function columnToAttribute(array $rawData, array $attributes, int $filterByKey = 0): array
    {
        $transformData = [];
        $attribute     = [];
        if (!isset($rawData[0]) || (\count($rawData[0]) !== \count($attributes))) {
            throw new \InvalidArgumentException('INVALID PARAMS');
        }
        
        foreach ($rawData as $key => $values) {
            if ($key === 0 || $rawData[$key][$filterByKey] === '') continue;
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
        return $this->columnToAttribute($rawValues, $attributes);
    }
    
    /**
     *  使用完成后释放内存
     */
    public function freeMemory()
    {
        $this->phpOffice->disconnectWorksheets();
        unset($this->phpOffice);
    }
    
    /**
     * 判断数组的深度
     *
     * @param $array
     * @return int
     */
    public static function arrayDepth(array $array): int
    {
        $maxDepth = 1;
        foreach ($array as $value) {
            if (\is_array($value)) {
                $depth = self::arrayDepth($value) + 1;
                if ($depth > $maxDepth) {
                    $maxDepth = $depth;
                }
            }
        }
        
        return $maxDepth;
    }
    
}