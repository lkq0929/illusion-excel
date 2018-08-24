importing and exporting Excel and CSV in yii2 
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


Usage
-----

Once the extension is installed, simply use it in your code by  :

```php
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
            'excel' => function() {
            return new \illusion\excel\Excel();
           },
       ],
    ****************************************/
    
    
    /**
     * 直接下载导出数据
     *
     * @return bool
     * ____________________________________________
     * notes: 
     * 将传入的数据填充到对应的单元格
     * export(array $cellData = [])
     * 
     * @param $cellData 组织好的要导出的数据，demo
     * [
     *  ['部门', '组别', '姓名', '性别'], //列标题 ,注意列标题与数据的对应
     *  ['一米技术部', 'oms', 'illusion', '男'],  //row1
     *  ['一米技术部', 'oms', 'alex', '男'],  //row2
     *  ['一米技术部', 'pms', 'aaron', '女'],
     * ]
     * return Spreadsheet
     *____________________________________________
     * 将传入的数据填充到对应的单元格
     * download($spreadSheet, string $fullFileName)
     * 
     * @param $fillData 填充数据后的Spreadsheet实例
     * @param $fullFileName 自定义文件名,例：export.xls
     * 直接在浏览器中下载
     * ____________________________________________
     * 
     * excel文件保存到本地
     * store($spreadSheet,  'test.xls')
     * @param $fillData 填充数据后的Spreadsheet实例
     * @param $fullFileName 自定义文件名,例：export.xls
     * ____________________________________________
     *
     */
    public function actionExport(): bool
    {
        $cellData = [
            ['部门', '组别', '姓名', '性别'],
            ['技术部', 'oms', 'illusion', '男'],
            ['技术部', 'oms', 'alex', '男'],
            ['技术部', 'pms', 'aaron', '女'],
        ];
        $fillData = \Yii::$app->excel->export($cellData);
        \Yii::$app->excel->download($fillData, 'test.xls');//直接下载导出excel
        //**\Yii::$app->excel->store($fillData,  'test.xls'); //默认导出excel保存到跟目录下,例：YII框架的web目录
        
        return true;
    }
    
    /**
     * 数据导入
     * @return mixed
     * ____________________________________________
     * notes:
     * import(string $tempPath, array $attributes = [], string $isRaw = '')
     * @param $tempPath 上传的excel文件的暂存路径
     * @param $attributes excel文件中的列按顺序对应的数据表属性
     * @param $isRaw   默认空字符串, 取值 'raw'或者''
     * return $data  , $isRaw = 'raw'的时候，只会返回文件中的原始数据,$isRaw空字符串的时候，返回传入属性格式化后的数据
     *____________________________________________
     * 
     */
    public function actionImport(): bool
    {
        $importFile = UploadedFile::getInstanceByName('file');
        $data       = \Yii::$app->excel->import($importFile->tempName, ['belongTo', 'group', 'name', 'sex'], 'raw');
        //$data 文件中处理后的数据，可以以你喜欢的方式批量或者循环插入数据表（这里推荐批量插入）
        
        return true;
    }
}

```
