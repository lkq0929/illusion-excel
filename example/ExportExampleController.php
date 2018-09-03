<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/9/3
 * Time: 14:42
 */

namespace app\controllers;


use yii\web\Controller;

class ExportExampleController extends Controller
{
    /**
     *  导出下载或者保存到根目录下的单工作表的电子表格文件
     */
    public function actionExportSheet()
    {
        $rows        = [
            ['部门', '组别', '姓名', '性别'],
            ['一米技术部', 'oms', 'illusion', '男'],
            ['一米技术部', 'oms', 'alex', '男'],
            ['一米技术部', 'pms', 'aaron', '女']
        ];
        $spreadsheet = new Excel();
        $sheet       = $spreadsheet->createSpreadSheet('xls');  //xls  电子表格文件类型
        $sheet->write($rows);
        $sheet->download('department');  //department 电子表格文件名
        //$sheetOne->store('department', './department');  //'department' 电子表格文件名  './department'  文件保存路径(实际保存位置：web/department/department.xls)
    }
    
    /**
     *  导出下载或者保存到根目录下的多工作表的电子表格文件
     */
    public function actionExportSheets()
    {
        $sheetRows = [
            'one' => [
                ['部门', '组别', '姓名', '性别'],
                ['一米技术部', 'oms', 'illusion', '男'],
                ['一米技术部', 'oms', 'alex', '男'],
                ['一米技术部', 'pms', 'aaron', '女']
            ],
            'two' => [
                ['类别', '名称', '价格'],
                ['文学类', '读者', '￥5'],
                ['科技类', 'AI之人工智能', '￥100'],
                ['科技类', '物联网起源', '￥500']
            ],
        ];
        
        $spreadsheet = new Excel();
        $sheets      = $spreadsheet->createSpreadSheet('xls');  //xls  电子表格文件类型
        $sheets->write($sheetRows);
        $sheets->download('department');  //department 电子表格文件名
        //$sheetOne->store('department', './department');  //'department' 电子表格文件名  './department'  文件保存路径(实际保存位置：web/department/department.xls)
    }
}