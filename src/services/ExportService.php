<?php

namespace illusion\excel\services;

class ExportService
{
    /**
     * 根据一行的个数生成excel对应的title元素个数
     *
     * EXAMPLE: $number = 27,  可以生成 ['A'....'Z', 'AA']
     *
     * auth: lkqlink@163.com
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
     * auth: lkqlink@163.com
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
    
}