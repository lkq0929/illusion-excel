<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/8/28
 * Time: 14:13
 */
namespace illusion\excel;

use yii\base\Component;

class Excel extends Component
{
    protected $typeList;
    
    public function __construct()
    {
        parent::__construct();
        $this->typeList = [
            'csv'  => __NAMESPACE__ . '\Csv',
            'xls'  => __NAMESPACE__ . '\Xls',
            'xlsx' => __NAMESPACE__ . '\Xlsx',
        ];
    }
    
    /**
     * 根据传入类型来创建相应的电子表格对象
     *
     * @param string $type
     * @return mixed
     * @throws NotSupportedException
     */
    public function createSpreadSheet(string $type)
    {
        if (!array_key_exists(strtolower($type), $this->typeList)) {
            throw new NotSupportedException();
        }
        $className = new $this->typeList[$type];
        
        return $className;
    }
}