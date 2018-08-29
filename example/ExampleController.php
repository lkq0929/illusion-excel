<?php
/**
 * Created by PhpStorm.
 *
 * auth: lkqlink@163.com
 * Date: 2018/8/23
 * Time: 15:35
 */
namespace app\controllers;


use yii\rest\Controller;

class ExampleController extends Controller
{
    
    /*****************配置说明****************
     * 'components' => [
     *      'excel' => [
     *          'class' => 'illusion\excel\Excel',
     *      ],
     * ],
     ****************************************/
    
    
    /**
     * 导出数据下载或者保存数据到服务器
     *
     * @return bool
     */
    public function actionExport(): bool
    {
        $cellData    = [
            ['部门', '组别', '姓名', '性别'],
            ['一米技术部', 'oms', 'illusion', '男'],
            ['一米技术部', 'oms', 'alex', '男'],
            ['一米技术部', 'pms', 'aaron', '女'],
        ];
        $spreadSheet = new Excel();
        $sheetOne    = $spreadSheet->createSpreadSheet('xls'); //实例化相应类型的电子表格实例
        $sheetOne->write($cellData);
        $sheetOne->download('oneSheetTwo'); //直接在浏览器中下载
        //$sheetOne->store('oneSheetTwo');  //保存到服务器
    }
    
    /**
     * 数据导入
     *
     * @return bool
     */
    public function actionImport(): bool
    {
        $importFile  = UploadedFile::getInstanceByName('file');
        $fileName    = explode('.', $importFile->name);
        $spreadSheet = new Excel();
        $sheetOne    = $spreadSheet->createSpreadSheet($fileName[1]); //实例化相应类型的电子表格实例
        $rawData     = $sheetOne->read($importFile->tempName); //读取excel文件中数据
        $data        = $sheetOne->columnTo($rawData, ['belongTo', 'group', 'name', 'sex']); //格式化原始数据和数据表中相应字段，注意与原始数据字段顺序的一致性
        /*example:
        原始数据:
        $rawData = [
            ['部门', '组别', '姓名', '性别'],
            ['一米技术部', 'oms', 'illusion', '男'],
            ['一米技术部', 'oms', 'alex', '男'],
            ['一米技术部', 'pms', 'aaron', '女'],
        ];
        columnTo函数转换后的数据:
        [
            {
                "belongTo": "一米技术部",
                "group": "oms",
                "name": "illusion",
                "sex": "男"
            },
            {
                "belongTo": "一米技术部",
                "group": "oms",
                "name": "alex",
                "sex": "男"
            },
            {
                "belongTo": "一米技术部",
                "group": "pms",
                "name": "aaron",
                "sex": "女"
            }
        ]*/
      
        return $data;
        //$data 文件中处理后的数据，可以以你喜欢的方式批量或者循环插入数据表（这里推荐批量插入）
    }
}