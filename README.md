PHPExcel for Kohanaframework
========

PHPExcel - OpenXML - Read, Write and Create Excel documents in PHP - Spreadsheet engine.

Include module:
1. Add to bootstrap.php:

```php
Kohana::modules(
    array(
        ...
        // Work with Excel
        'phpexcel' => MODPATH.'phpexcel',
        ...
    )
);
```

2. Load new version of PHPExcel from http://phpexcel.codeplex.com/

How use:

1. Without pages in xls document:

```php
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
```

2. Has many pages:

```php
$data = array(
    'page1' => array(
        1 => // строка
            array(
                'ceil1',
                'ceil2',
                'ceil3'
            )
    ),
    'page2' => array(
        1 => // строка
            array(
                'ceil1',
                'ceil2',
                'ceil3'
            )
    ),
);
 
$XLSX = new Spreadsheet();
$XLSX->setData( $data, 1 );
$XLSX->load(array('name' => 'Excel_FileName')));

3. All features:

// with image on page
$iDrawing = new PHPExcel_Worksheet_Drawing();
 
// get image
$iDrawing->setPath('/images/path.png');
 
// set img ceil
$iDrawing->setCoordinates('A1');
 
// image offset X и Y
$iDrawing->setOffsetX(0);
$iDrawing->setOffsetY(0);
 
// if don't need save proportional sizes for image
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
    'subject' => 'subj',
    'description' => 'desc',
    'author' => 'author name',
));
 
// Fill ceil styles (background)
$grey_fill_style = array(
    'fill' => array(
        'type'     => PHPExcel_Style_Fill::FILL_GRADIENT_PATH,
        'rotation'   => 0,
        'startcolor' => array(
            'rgb' => 'AAAAAA'
        ),
        'endcolor'   => array(
            'argb' => 'AAAAAA'
        )
    )
);
 
// the map of styles in ceils
$styles = array(
    1 => array(array(), array(), array(), array(), $grey_fill_style),
    2 => array(array('font' => array( // настройка шрифта в ячейке
        //'name'      => 'Arial',
        'bold'    => true,
        'italic'    => false,
        'underline' => false,
        'strike'    => false,
        'color'  => array(
            'rgb' => '000000'
        ),
        'size' => 18
    ))),
    3 => array($grey_fill_style, $grey_fill_style, $grey_fill_style,
        $grey_fill_style, $grey_fill_style),
);
 
for($i=4; $i<=20; $i++)
    $styles[$i] = array($grey_fill_style);
 
// alignment in ceil
$alignment =  array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::VERTICAL_JUSTIFY,
        'vertical'   => PHPExcel_Style_Alignment::VERTICAL_TOP,
        'rotation'   => 0,
        'wrap'     => true // обрабатывать \n в тексте
    )
);
 
$styles[6][1] = $alignment;
$styles[6][0] = Arr::merge($styles[6][0], $alignment);
 
// merge ceils
$merging = array(
    'A1:D1', 'A2:E2', 'A3:D3', 'B6:E6'
);
 
$XLSX->setData( $pages, true, $styles, $merging, true
    /** auto width of ceil */ );
 
$XLSX->load(array('name' => $title));
```