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
use yii\web\UploadedFile;

class ImportExampleController extends Controller
{
    /**
     * 单工作表的电子表导入
     *
     * @return array
     */
    public function actionImportSheet(): array
    {
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        
        $rawData   = \Yii::$app->excel->read($importFile->tempName);  //$importFile->tempName 文件路径  返回数据：电子表格里面的原生数据，可自己处理
        $transData = \Yii::$app->excel->rawValuesToAttribute($rawData, ['department', 'group', 'name', 'sex']);  //$rawData 单个工作表原生数据  ['department', 'group', 'name', 'sex'] 单个工作表每列对应的属性
        
        return $transData;
    }
    
    /**
     * 多工作表的电子表导入
     *
     * @return array
     */
    public function actionImportSheets(): array
    {
        $transData  = [];
        $attributes = [
            'one' => ['department', 'group', 'name', 'sex'],
            'two' => ['category', 'book_name', 'price'],
        ];
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        $rawDatas = \Yii::$app->excel->read($importFile->tempName);
        foreach ($rawDatas as $sheetName => $rawData) {
            $transData[] = \Yii::$app->excel->rawValuesToAttribute($rawData, $attributes[$sheetName]);
        }
        
        return $transData;
    }
    
    /**
     * .csv格式的文件只能工作在（单sheet）模式下
     * csv后缀的excel文件处理
     *
     * @return mixed
     */
    public function actionImportCsv()
    {
        $spreadSheet = \Yii::$app->excel;
        $customAttr  = ['name', 'sex', 'age']; //自定义属性
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        list($fileName, $ext) = explode('.', $importFile->name);
        $spreadSheet->setSuffix($ext);  //设置电子表格后缀
        //将把自定义的属性映射到工作表对应的列，方便直接后面直接插入数据表
        $rawData = $spreadSheet->rawValuesToAttribute($spreadSheet->read($importFile->tempName), $customAttr);
        
        return $rawData;
    }
    
}