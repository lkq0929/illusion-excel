<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/9/26
 * Time: 16:57
 */

namespace illusion\excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Spreadsheet as phpSpreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use yii\base\Component;

class Spreadsheet extends Component
{
    
    const SUFFIX_XLS = 'xls';
    const SUFFIX_CSV = 'csv';
    const SUFFIX_XLSX = 'xlsx';
    
    private static $supportSuffix = [
        self::SUFFIX_CSV,
        self::SUFFIX_XLS,
        self::SUFFIX_XLSX,
    ];
    
    /**
     * @var 文件名称
     */
    private $name;
    
    /**
     * @var 文件后缀
     */
    private $suffix = 'xls';
    
    /**
     * @var 工作表
     */
    private $worksheet;
    
    protected $_spreadsheet;
    
    public function __construct(array $config = [])
    {
        parent::__construct($config);
        $this->_spreadsheet = new phpSpreadsheet();
        $this->worksheet    = new Worksheet($this->_spreadsheet);
    }
    
    /**
     * 设置电子表格名称
     *
     * @param $name
     */
    public function setSpreadsheetName($name)
    {
        $this->name = $name;
    }
    
    /**
     * 设置电子表格后缀
     *
     * @param $suffix
     */
    public function setSuffix($suffix)
    {
        $this->suffix = strtolower($suffix);
    }
    
    /**
     * 设置下载头部信息
     */
    public function setHeader()
    {
        $mime = $this->getHeaderMime($this->suffix);
        \Yii::$app->response->setDownloadHeaders($this->name . '.' . $this->suffix, $mime)->send();
    }
    
    /**
     * 获取浏览器传输内容类型
     *
     * @param string $suffix
     * @return string
     */
    public function getHeaderMime($suffix): string
    {
        $mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        if (\in_array($suffix, [static::SUFFIX_XLS, static::SUFFIX_CSV], true)) {
            $mime = 'application/vnd.ms-excel';
        }
        
        return $mime;
    }
    
    /**
     * 获取电子表格名称
     *
     * @return 文件名称
     */
    public function getSpreadsheetName()
    {
        return $this->name;
    }
    
    /**
     * 是否支持该后缀
     *
     * @param string $suffix
     * @return bool
     */
    public static function isSupportSuffix(string $suffix): bool
    {
        $isSupport = false;
        if (\in_array($suffix, self::$supportSuffix, true)) {
            $isSupport = true;
        }
        
        return $isSupport;
    }
    
    /**
     * 写对对应后缀的电子表格文件
     *
     * @param string $suffix
     * @return IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writerToType($suffix): IWriter
    {
        return IOFactory::createWriter($this->_spreadsheet, ucfirst($suffix));
    }
    
    /**
     * 计算数据的维度
     *
     * @param array $array
     * @return int
     */
    public static function arrayDepth(array $array): int
    {
        $maxDepth = 1;
        foreach ($array as $value) {
            if (\is_array($value)) {
                $depth = self::arrayDepth($value) + 1;
                if ($depth > $maxDepth) {
                    $maxDepth = $depth;
                }
            }
        }
        
        return $maxDepth;
    }
    
    /**
     * 写入数据到电子表
     *
     * @param array $values
     * @return phpSpreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function write(array $values): phpSpreadsheet
    {
        $valuesDepth = static::arrayDepth($values);
        if ($valuesDepth <= 1 || $valuesDepth > 3) {
            throw new \InvalidArgumentException('INVALID　PARAMS');
        }
        
        if ($valuesDepth === 2) {
            $values = [$values];
        }
        
        return $this->worksheet->fillValuesToSheets($values);
    }
    
    /**
     * 浏览器直接下载电子表
     *
     * @param $spreadsheetName
     * @param $suffix
     * @throws NotSupportedException
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function download($spreadsheetName, $suffix)
    {
        if (!static::isSupportSuffix($suffix)) {
            throw new NotSupportedException("NOT SUPPORTED {$suffix} TYPE");
        }
        $this->setSpreadsheetName($spreadsheetName);
        $this->setSuffix($suffix);
        $this->setHeader();
        $this->writerToType($this->suffix)->save('php://output');
    }
    
    /**
     * 保存电子表到服务器，Yii框架默认下载到web目录下，保存路径可自定义
     *
     * @param $spreadsheetName
     * @param $suffix
     * @param string $path
     * @throws NotSupportedException
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function store($spreadsheetName, $suffix, string $path = './')
    {
        if (!static::isSupportSuffix($suffix)) {
            throw new NotSupportedException("NOT SUPPORTED {$suffix} TYPE");
        }
        $this->setSpreadsheetName($spreadsheetName);
        $this->setSuffix($suffix);
        $this->writerToType($this->suffix)->save($path . $this->name . '.' . $this->suffix);
    }
    
    /**
     * 读取电子表格中所有工作表的数据
     *
     * @param $path
     * @return array
     * @throws NotSupportedException
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function read($path): array
    {
        if (!static::isSupportSuffix($this->suffix)) {
            throw new NotSupportedException("NOT SUPPORTED {$this->suffix} TYPE");
        }
        if ($this->suffix === 'csv') {
            return $this->readForCsv($path);
        }
        $spreadSheet = IOFactory::load($path);
    
        return $this->getSheetsValues($spreadSheet);
    }
    
    /**
     * csv文件数据读取
     *
     * @param $path
     * @return array
     */
    public function readForCsv($path): array
    {
        
        $file = fopen($path, "r");
        while (!feof($file)) {
            $data[] = fgetcsv($file);
        }
        $data = eval('return ' . iconv('gbk', 'utf-8', var_export($data, true)) . ';');
        foreach ($data as $key => $value) {
            if (!$value) {
                unset($data[$key]);
            }
        }
        fclose($file);
    
        return $data;
    }
    
    /**
     *
     *
     * @param $spreadSheet
     * @return array
     */
    public function getSheetsValues($spreadSheet): array
    {
        $rawValues   = $this->worksheet->getSheetsValues($spreadSheet);
    
        return $rawValues;
    }
    
    /**
     * 工作表中的数据格式转换为数据库数据表中的格式
     *
     * @param array $rawValues
     * @param array $attributes
     * @return array
     */
    public function rawValuesToAttribute(array $rawValues = [], array $attributes = []): array
    {
        return $this->columnToAttribute($rawValues, $attributes);
    }
    
    /**
     * 工作表类属性到数据库数据表属性名称的映射
     *
     * @param array $rawData
     * @param array $attributes
     * @param int $filterByKey
     * @return array
     */
    public function columnToAttribute(array $rawData, array $attributes, int $filterByKey = 0): array
    {
        $transformData = [];
        $attribute     = [];
        if (!isset($rawData[0]) || (\count($rawData[0]) !== \count($attributes))) {
            throw new \InvalidArgumentException('INVALID PARAMS');
        }
        
        foreach ($rawData as $key => $values) {
            if ($key === 0 || $rawData[$key][$filterByKey] === '') continue;
            foreach ($values as $column => $value) {
                $attribute[$attributes[$column]] = $value;
            }
            $transformData[] = $attribute;
        }
        
        return $transformData;
    }
    
    
    /**
     * 释放内存
     */
    public function __destruct()
    {
        $this->_spreadsheet->disconnectWorksheets();
        unset($this->_spreadsheet);
    }
}