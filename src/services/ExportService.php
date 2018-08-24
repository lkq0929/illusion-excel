<?php
/**
 * Created by PhpStorm.
 *
 * auth: lkqlink@163.com
 * Date: 2018/8/7
 * Time: 14:25
 */
namespace illusion\excel\services;

use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class ExportService
{
    /**
     * 根据一行的个数生成excel对应的title元素个数
     *
     * EXAMPLE: $number = 27,  可以生成 ['A'....'Z', 'AA']
     *
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
     * 获取不同格式对应的浏览器内容输出类型
     *
     * @param string $ext
     * @return string
     */
    public static function getExtToMime(string $ext): string
    {
        $transformExt = ucfirst($ext);
        $mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        if (\in_array($transformExt, ['Xls', 'Csv'])) {
            $mime = 'application/vnd.ms-excel';
        }
    
        return $mime;
    }
    
    /**
     * 处理文件名
     *
     * @param string $fullFileName
     * @return array
     * @throws Exception
     */
    public static function handleFileName(string $fullFileName): array
    {
        self::validate($fullFileName);
        list($result['fileName'], $result['ext']) = explode('.', $fullFileName);
        
        return $result;
    }
    
    /**
     * 数据写入excel文件
     *
     * @param Spreadsheet $spreadsheet
     * @param string $ext
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public static function writeData(Spreadsheet $spreadsheet, string $ext): IWriter
    {
        $exportFile = IOFactory::createWriter($spreadsheet, ucfirst($ext));
        
        return $exportFile;
    }
    
    /**
     * 文件名的验证
     *
     * @param $attribute
     * @return bool
     * @throws Exception
     */
    public static function validate(string $attribute): bool
    {
        if (!preg_match("/(\.xls|\.xlsx|\.csv)/", strtolower($attribute)) || substr_count($attribute, '.') !== 1
            || strpos($attribute, '.') === 0) {
            throw new Exception('不合法的文件名');
        }
    
        return true;
    }
    
}