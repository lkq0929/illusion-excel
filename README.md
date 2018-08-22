# illusion-excel
excel import and export for yii2

## Installation

```bash
$> composer require illusion/yii2-excel
```

## A simple example

```php
<?php

namespace app\controllers;

use yii\rest\Controller;

class Example extends Controller
{
    //浏览器直接下载导出的数据
    public function actionTestExport()
    {
        $cellData = [
            ['部门', '组别', '姓名', '性别'],
            ['一米技术部', 'oms', 'illusion', '男'],
            ['一米技术部', 'oms', 'alex', '男'],
            ['一米技术部', 'pms', 'aaron', '女'],
        ];
        $excel = new excel\Excel();
        $fillData = $excel->export($cellData);
        $excel->download($fillData, 'test', 'Xls'); 
        
        return true;
    }
    //导入的数据进行格式化
    public function actionTestImport()
    {
        $importFile = UploadedFile::getInstanceByName('file');
        $excel = new excel\Excel($importFile->tempName);
        $data = $excel->import(['belongTo', 'group', 'name', 'sex'], 'transform'); //['belongTo', 'group', 'name', 'sex']，excel文件中对应的数据表属性，注意表的属性顺序应与excel文件中列按照顺序一一对应
       
        return $data;  //最后直接将装换后的数组批量插入即可
    }
}
```
