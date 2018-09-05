excel import and export for yii2
================================
excel import and export for yii2

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist illusion/yii2-excel "*"
下载一个具体的版本：
example: composer require illusion/yii2-excel 1.2.0
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
            'class' => 'illusion\excel\Excel',
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
     * 单电子表(sheet)、多电子表格导出
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
        $spreadsheet    = \Yii::$app->excel->createSpreadSheet('xls'); // 'xls' 自定义电子表格后缀
        $spreadsheet->write($cellData);
        $spreadsheet->download('department'); //'department' 自定义电子表格名
    }
    
    /**
     * 读出电子表格的原生数据并根据自定义属性名装换成对应格式
     * 然后根据你喜欢的方式将数据插入数据表
     *
     * @return array
     */
    public function actionImport()
    {
        $transData  = [];
        $attributes = [
            'department' => ['department', 'group', 'name', 'sex'],  //'department' 工作表(sheet)名  键值：['department', 'group', 'name', 'sex']列对应的名称,名称顺序必须一致
            'book'       => ['category', 'book_name', 'price'],
        ];
        if (\Yii::$app->request->isPost) {
            $importFile = UploadedFile::getInstanceByName('file');
        }
        $fileName    = explode('.', $importFile->name);
        $spreadsheet = \Yii::$app->excel->createSpreadSheet($fileName[1]); // $fileName[1] 自定义电子表格后缀
        $rawDatas    = $spreadsheet->read($importFile->tempName); //$importFile->tempName 电子表（excel）路径
        foreach ($rawDatas as $sheetName => $rawData) {
            $transData[] = $spreadsheet->columnTo($rawData, $attributes[$sheetName]);  //$rawData 读出的单个工作表(sheet)的原生数据，$attributes[$sheetName] $rawData电子表格中列按照顺序对应的自定义列名
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
        $fileName    = explode('.', $importFile->name);
        $spreadsheet = \Yii::$app->excel->createSpreadSheet($fileName[1]); //同上注释
        $rawDatas    = $spreadsheet->read($importFile->tempName);
        
        return $rawDatas;
    }
}
```
