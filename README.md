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
            'class' => 'illusion\excel\Excel',
        ],
    ],
];
`````
demo
-----
```
https://github.com/lkq0929/illusion-excel/tree/master/example
```
