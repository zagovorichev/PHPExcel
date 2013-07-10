# Скопировать phpexcel в директорию *MODPATH*

## Структура модуля:

![Структура каталогов](structure.jpg)

# В bootstrap.php подключить новый модуль
~~~
Kohana::modules(
    array(
        ...
        // Work with Excel
        'phpexcel' => MODPATH.'phpexcel',
        ...
    )
);
~~~

# Загрузить новую версию PHPExcel
в директорию *MODPATH/phpexcel/vendor* из [официального источника](http://phpexcel.codeplex.com/ "PHPExcel Codeplex")