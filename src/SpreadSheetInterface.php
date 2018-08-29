<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/8/28
 * Time: 13:52
 */
namespace illusion\excel;

interface SpreadSheetInterface {
    /**
     * 写数据
     *
     * @param $data
     * @return mixed
     */
    public function write(array $data);
    
    /**
     * 读数据
     *
     * @param $path
     * @return mixed
     */
    public function read(string $path);
}