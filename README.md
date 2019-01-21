excel import and export for yii2
================================
excel import and export for yii2

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist illusion/yii2-excel "*"
```

or add

```
"illusion/yii2-excel": "*"
```

to the require section of your `composer.json` file.


config
-----

Once the extension is installed, simply use it in your code by  :

```php
return [
    'components' => [
        'excel' => [
            'class' => 'illusion\excel\Spreadsheet',
        ],
    ],
];
`````
EXAMPLE LINK
-----
```
https://github.com/lkq0929/illusion-excel/tree/master/example
```
```php
<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/8/27
 * Time: 14:28
 */

namespace app\controllers;


use yii\rest\Controller;
use yii\web\UploadedFile;

class ExampleController extends Controller
{
    /**
     * 单工作表(sheet)、多工作表的电子表格导出
     * 导出文件可直接浏览器下载或保存至服务器
     */
    public function actionExport()
    {
        /*单电子表的电子表格--数据格式
        $cellData = [
            ['部门', '组别', '姓名', '性别'],
            ['一米技术部', 'oms', 'illusion', '男'],
            ['一米技术部', 'oms', 'alex', '男'],
            ['一米技术部', 'pms', 'aaron', '女']
        ];*/
        /*多电子表的电子表格--数据格式*/
        $cellData = [
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
        \Yii::$app->excel->write($cellData);
        \Yii::$app->excel->download('department','xls'); //'department' 自定义电子表格名,直接下载名称为department.xls的文件
        //\Yii::$app->excel->store('department','xls','./'); //'department' 自定义电子表格名,保存名称为department.xls
    }
    
    /**
     * 读出电子表格的原生数据并根据自定义属性名装换成对应格式
     * 然后根据你喜欢的方式将数据插入数据表
     *
     * @return array
     */
    public function actionImport()
    {
        $transData   = [];
        $attributes  = [
            'one' => ['department', 'group', 'name', 'sex'],
            'two' => ['category', 'book_name', 'price'],
        ];
        if (\Yii::$app->request->isPost) {
            $importFile  = UploadedFile::getInstanceByName('file');
        }
        $rawDatas    = \Yii::$app->excel->read($importFile->tempName);
        foreach($rawDatas as  $sheetName => $rawData) {
            $transData[] = \Yii::$app->excel->rawValuesToAttribute($rawData, $attributes[$sheetName]);
        }
        
        return $transData;
    }
    
    /**
     * 读出电子表格的原生数据
     * 根据喜欢的方式插入数据表
     *
     * @return mixed
     */
    public function actionImportRaw()
    {
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        $rawDatas  = \Yii::$app->excel->read($importFile->tempName);

        return $rawDatas;
    }   
    
    /**********************后缀csv格式的电子表格的两种读取方式，推荐方式一************************/
    /*注意csv格式的电子表格只能以单工作表模式（也就是一个文件一个sheet，不支持多sheet工作模式）*/
    
    /**
     *  方式一：
     *  **注意设置电子表格后缀
     * 将把自定义的属性（['name', 'sex', 'age']）映射到工作表对应的列，方便直接后面直接插入数据表
     * 根据喜欢的方式插入数据表
     *
     * @return mixed
     */
    public function actionImportCsvToAttr()
    {
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        $spreadsheet = \Yii::$app->excel;
        list($fileName, $ext) = explode('.', $importFile->name);
        $spreadsheet->setSuffix($ext);  //设置电子表格后缀
        $rawData = $spreadsheet->rawValuesToAttribute($spreadsheet->read($importFile->tempName), ['name', 'sex', 'age']); //单工作表格：将自定义的列名称映射到对应的电子表格中的列名称

        return $rawData;
    }
    /**
     * 方式二：
     * 读出csv格式电子表格的原生数据
     * 根据喜欢的方式插入数据表
     *
     * @return mixed
     */
    public function actionImportCsvRaw()
    {
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        $rawDatas  = \Yii::$app->excel->readForCsv($importFile->tempName);

        return $rawDatas;
    }
    
    /**
     * 读出csv格式电子表格的原生数据
     * 根据喜欢的方式插入数据表
     *
     * @return mixed
     */
    public function actionImportCsvRaw()
    {
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        $rawDatas  = \Yii::$app->excel->readForCsv($importFile->tempName);

        return $rawDatas;
    }
}
```
