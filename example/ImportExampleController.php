<?php
/**
     * Created by PhpStorm.
     *
     * Auth: lkqlink@163.com
     * Date: 2018/9/3
     * Time: 15:03
     */
    
namespace app\controllers;


use yii\web\Controller;

class ImportExampleController extends Controller
{
    /**
     * 单工作表的电子表导入
     *
     * @return array
     */
    public function actionImportSheet():array
    {
        $spreadSheet = new Excel();
        $importFile  = UploadedFile::getInstanceByName('file');
        $fileName    = explode('.', $importFile->name);
    
        $sheetOne  = $spreadSheet->createSpreadSheet($fileName[1]);  //$fileName[1]  电子表格文件类型
        $rawData   = $sheetOne->read($importFile->tempName);  //$importFile->tempName 文件路径  返回数据：电子表格里面的原生数据，可自己处理
        $transData = $sheetOne->columnTo($rawData, ['department', 'group', 'name', 'sex']);  //$rawData 单个工作表原生数据  ['department', 'group', 'name', 'sex'] 单个工作表每列对应的属性
    
        return $transData;
    }
    
    /**
     * 多工作表的电子表导入
     *
     * @return array
     */
    public function actionImportSheets(): array
    {
        $transData   = [];
        $spreadSheet = new Excel();
        $attributes  = [
            'one' => ['department', 'group', 'name', 'sex'],
            'two' => ['category', 'book_name', 'price'],
        ];
        $importFile  = UploadedFile::getInstanceByName('file');
        $fileName    = explode('.', $importFile->name);
        $sheetOne    = $spreadSheet->createSpreadSheet($fileName[1]);  //$fileName[1]  电子表格文件类型
        $rawDatas    = $sheetOne->read($importFile->tempName);  //$importFile->tempName 文件路径  返回数据：电子表格里面的原生数据，可自己处理
        foreach($rawDatas as  $sheetName => $rawData) {
            $transData[] = $sheetOne->columnTo($rawData, $attributes[$sheetName]);  //$rawData 单个工作表原生数据  $attributes[$sheetName] 单个工作表每列对应的属性
        }
    
        return $transData;
    }
    
}