<?php
/**
     * Created by PhpStorm.
     *
     * Auth: lkqlink@163.com
     * Date: 2018/8/28
     * Time: 14:00
     */
    
namespace illusion\excel;


class Xls extends  BaseSpreadSheet implements SpreadSheetInterface
{
    public function __construct()
    {
        parent::__construct();
    }
    
    /**
     * 下载xls类型的文件
     *
     * @param string $spreadSheetName
     * @param string $ext
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function download(string $spreadSheetName, string $ext = 'xls')
    {
        $this->setHeader($spreadSheetName, $ext);
        $this->writerToType($ext)->save('php://output');
    }
    
    /**
     * 保存xls格式的电子表格到自定义路径下
     *
     * @param string $path
     * @param string $spreadSheetName
     * @param string $ext
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function store(string $spreadSheetName, string $path = './', string $ext = 'xls')
    {
        $this->writerToType($ext)->save($path . $spreadSheetName . '.' . strtolower($ext));
    }
    
}