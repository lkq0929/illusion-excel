<?php
/**
 * Created by PhpStorm.
 *
 * Auth: lkqlink@163.com
 * Date: 2018/8/27
 * Time: 13:44
 */
    
namespace illusion\excel;


use Throwable;

class NotSupportedException extends \Exception
{
    /**
     * 不支持功能异常
     *
     * @return string
     */
    public function __construct(string $message = 'Not Supported', int $code = 0, Throwable $previous = null)
    {
        parent::__construct($message, $code, $previous);
    }
    
    
}