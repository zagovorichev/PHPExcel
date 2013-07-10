# Использование подуля

## Без страниц в excel документе
~~~
$data = array(
    array(
        1 => // строка
            array(
                'колонка 1',
                'колонка 2',
                'колонка 3'
            )
    ),
);


$XLSX = new Spreadsheet();
$XLSX->setData( $data );
$XLSX->load(array('name' => 'Excel_FileName')));
~~~

## С использованием страниц

~~~
$data = array(
    'Страница 1' => array(
        1 => // строка
            array(
                'колонка 1',
                'колонка 2',
                'колонка 3'
            )
    ),
    'Страница 2' => array(
        1 => // строка
            array(
                'колонка 1',
                'колонка 2',
                'колонка 3'
            )
    ),
);


$XLSX = new Spreadsheet();
$XLSX->setData( $data, 1 );
$XLSX->load(array('name' => 'Excel_FileName')));
~~~

## Использование всех возможностей
~~~
$iDrawing = new PHPExcel_Worksheet_Drawing();

//берем рисунок
$iDrawing->setPath('/images/path.png');

//устанавливаем ячейку
$iDrawing->setCoordinates('A1');

//устанавливаем смещение X и Y
$iDrawing->setOffsetX(0);
$iDrawing->setOffsetY(0);

//если не нужно сохранять пропорции изображения
//$iDrawing->setResizeProportional(false);
$iDrawing->setHeight(50);

$pages = array(
    'page1_name' => array(
        1 => array(clone $iDrawing, '', '', '', 'ceil5'),

        2 => array('some text for ceil1'),

        3 => array('text', '','','', 'ceil5' /** ... ceil N */),
        //... row N
    ),
    // ... page N
);

$XLSX = new Spreadsheet(array(
    'title' => 'doc_title',
    'subject' => 'Тема',
    'description' => 'Описание документа',
    'author' => 'Автор документа',
));

// Стили для заполнения ячеек (background)
$grey_fill_style = array(
    'fill' => array(
        'type'	   => PHPExcel_Style_Fill::FILL_GRADIENT_PATH,
        'rotation'   => 0,
        'startcolor' => array(
            'rgb' => 'AAAAAA'
        ),
        'endcolor'   => array(
            'argb' => 'AAAAAA'
        )
    )
);

// Карта стилей для ячеек
$styles = array(
    1 => array(array(), array(), array(), array(), $grey_fill_style),
    2 => array(array('font' => array( // настройка шрифта в ячейке
        //'name'	  => 'Arial',
        'bold'	  => true,
        'italic'	=> false,
        'underline' => false,
        'strike'	=> false,
        'color'	 => array(
            'rgb' => '000000'
        ),
        'size' => 18
    ))),
    3 => array($grey_fill_style, $grey_fill_style, $grey_fill_style,
        $grey_fill_style, $grey_fill_style),
);

for($i=4; $i<=20; $i++)
    $styles[$i] = array($grey_fill_style);

// Выравнивание в ячейке
$alignment =  array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::VERTICAL_JUSTIFY,
        'vertical'   => PHPExcel_Style_Alignment::VERTICAL_TOP,
        'rotation'   => 0,
        'wrap'	   => true // обрабатывать \n в тексте
    )
);

$styles[6][1] = $alignment;
$styles[6][0] = Arr::merge($styles[6][0], $alignment);

// объединяемые ячейки на всех страницах документа
$merging = array(
    'A1:D1', 'A2:E2', 'A3:D3', 'B6:E6'
);

$XLSX->setData( $pages, true, $styles, $merging, true
    /** Растягивать столбцы по содержимому (auto width) */ );

$XLSX->load(array('name' => $title));
~~~
