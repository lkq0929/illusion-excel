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
     * auth: lkqlink@163.com
     * ____________________________________________
     * notes:
     * 将传入的数据填充到对应的单元格
     * export(array $cellData = [])
     *
     * @param $cellData 将
     * return Spreadsheet
     *____________________________________________
     * 将传入的数据填充到对应的单元格
     * download($spreadSheet, string $fileName, string $ext)
     *
     * @param $fillData 填充数据后的Spreadsheet实例
     * @param $fileName 导出的excel文件名
     * @param $ext  导出的excel文件名后缀
     * 直接在浏览器中下载
     * ____________________________________________
     * @return bool
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
        \Yii::$app->excel->download($fillData, 'test', 'Xls');
        
        return true;
    }
    
    /**
     * 数据导入
     *
     * auth: lkqlink@163.com
     * ____________________________________________
     * notes:
     * import(string $tempPath, array $attributes = [], string $isRaw = '')
     * @param $tempPath 上传的excel文件的暂存路径
     * @param $attributes excel文件中的列按顺序对应的数据表属性
     * @param $isRaw   默认空字符串, 取值 'raw'或者''
     * return $data  , $isRaw = 'raw'的时候，只会返回文件中的原始数据,$isRaw空字符串的时候，返回传入属性格式化后的数据
     *____________________________________________
     * @return mixed
     */
    public function actionImport(): bool
    {
        $importFile = UploadedFile::getInstanceByName('file');
        $data       = \Yii::$app->excel->import($importFile->tempName, ['belongTo', 'group', 'name', 'sex'], 'raw');
        //$data 文件中处理后的数据，可以以你喜欢的方式批量或者循环插入数据表（这里推荐批量插入）
        
        return true;
    }
}