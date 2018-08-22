<?php
/**
 * Created by PhpStorm.
 *
 * Date: 2018/8/22
 * Time: 11:11
 */
    
namespace illusion\excel\services;


class ImportService
{
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
        foreach ($rawData as $key => $values) {
                if ($key === 0) continue ;
                foreach ($values as $column => $value) {
                    $attribute[$attributes[$column]]= $value;
                }
                $transformData[] = $attribute;
        }
        
        return $transformData;
    }

}