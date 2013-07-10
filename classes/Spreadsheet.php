<?php defined('SYSPATH') or die('No direct access allowed.');
/**
 * PHP Excel library. Helper class to make spreadsheet creation easier.
 *
 * @package    Spreadsheet
 * @author     Flynsarmy
 * @website    http://www.flynsarmy.com/
 * @license    TEH FREEZ
 */
class Spreadsheet
{
	const VENDOR_PACKAGE = "vendor/PHPExcel/PHPExcel/";
	private $_spreadsheet;

	/*
	 * Purpose: Creates the spreadsheet with given or default settings
	 * Input: array $headers with optional parameters: title, subject, description, author
	 * Returns: void
	 */
	public function __construct($headers=array())
	{
		$headers = array_merge(array(
			'title'			=> 'New Spreadsheet',
			'subject'		=> 'New Spreadsheet',
			'description'	=> 'New Spreadsheet',
			'author'		=> 'ClubSuntory',

		), $headers);

		$this->_spreadsheet = new PHPExcel();
		// Set properties
		$this->_spreadsheet->getProperties()
			->setCreator( $headers['author'] )
			->setTitle( $headers['title'] )
			->setSubject( $headers['subject'] )
			->setDescription( $headers['description'] );
			//->setActiveSheetIndex(0);
		//$this->_spreadsheet->getActiveSheet()->setTitle('Minimalistic demo');
	}

	/*
	 * Purpose Writes cells to the spreadsheet
	 * Input: array of array( [row] => array([col]=>[value]) ) ie $arr[row][col] => value
	 * Returns: void
	 */
	public function setData(array $data, $multi_sheet=false, $styles=array(), $merging_ceil=array(), $auto_width=false)
	{
		if ( empty($this->_spreadsheet) )
			$this->create();

		//Single sheet ones can just dump everything to the current sheet
		if ( !$multi_sheet )
		{
			$Sheet = $this->_spreadsheet->getActiveSheet();
			$this->setSheetData( $data, $Sheet, $styles, $merging_ceil, $auto_width );
		}
		//Have to do a little more work with multi-sheet
		else
		{
			foreach ( $data as $sheetName=>$sheetData )
			{
				$Sheet = $this->_spreadsheet->createSheet();
				$Sheet->setTitle( $sheetName );
				$this->setSheetData( $sheetData, $Sheet, $styles, $merging_ceil, $auto_width );
			}
			//Now remove the auto-created blank sheet at start of XLS
			$this->_spreadsheet->removeSheetByIndex( 0 );
		}

		/*
		array(
			1 => array('A1', 'B1', 'C1', 'D1', 'E1')
			2 => array('A2', 'B2', 'C2', 'D2', 'E2')
			3 => array('A3', 'B3', 'C3', 'D3', 'E3')
		);
		*/
	}

	public function setSheetData( array $data, PHPExcel_Worksheet $Sheet, $styles = array(), $merging_ceil=array(), $auto_width=false )
	{
		foreach ( $data as $row => $columns )
			foreach ( $columns as $column => $value ){
                if(is_a($value, 'PHPExcel_Worksheet_Drawing')){
                    $value->setWorksheet($Sheet);
                    $Sheet->getRowDimension()->setRowHeight($value->getHeight());
                } else {
                    $Sheet->setCellValueByColumnAndRow($column, $row, $value);
                }

                if(count($styles) && isset($styles[$row][$column]) && count($styles[$row][$column]))
                    $Sheet->getStyleByColumnAndRow($column, $row)->applyFromArray($styles[$row][$column]);

                if(count($merging_ceil)) foreach($merging_ceil as $merge) $Sheet->mergeCells($merge);

                if(is_string($value)){

                    // set row height (if in one of cell have(\n))
                    if( strpos($value, "\n")!==false ){
                        $s_num = substr_count($value, "\n")+1;

                        $height = $Sheet->getStyleByColumnAndRow($column, $row)->getFont()->getSize()+5;

                        //def font style
                        $new_height = $height*$s_num;

                        $dimension = $Sheet->getRowDimension($row);
                        if($dimension->getRowHeight() < $new_height)
                            $dimension->setRowHeight($new_height);
                    }

                    //ширина по тексту
                    if($auto_width)
                        $Sheet->getColumnDimensionByColumn($column)->setAutoSize(true);

                }
            }
	}

	/*
	 * Purpose: Writes spreadsheet to file
	 * Input: array $settings with optional parameters: format, path, name (no extension)
	 * Returns: Path to spreadsheet
	 */
	public function save( $settings=array() )
	{
		if ( empty($this->_spreadsheet) )
			$this->create();

		//Used for saving sheets
		require self::VENDOR_PACKAGE.'IOFactory.php';

		$settings = array_merge(array(
			'format'		=> 'Excel2007',
			'path'			=> APPPATH.'assets/downloads/spreadsheets/',
			'name'			=> 'NewSpreadsheet'

		), $settings);

		//Generate full path
		$settings['fullpath'] = $settings['path'] . $settings['name'] . '_'.time().'.xlsx';

		$Writer = PHPExcel_IOFactory::createWriter($this->_spreadsheet, $settings['format']);
		// If you want to output e.g. a PDF file, simply do:
		//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');
		$Writer->save( $settings['fullpath'] );

		return $settings['fullpath'];
	}

    public function load( $settings=array() )
    {
        if ( empty($this->_spreadsheet) )
            $this->create();

        //Used for saving sheets
        require self::VENDOR_PACKAGE.'IOFactory.php';

        $settings = array_merge(array(
            'format'		=> 'Excel2007',
            'path'			=> APPPATH.'assets/downloads/spreadsheets/',
            'name'			=> 'NewSpreadsheet'

        ), $settings);

        $Writer = PHPExcel_IOFactory::createWriter($this->_spreadsheet, $settings['format']);
        // If you want to output e.g. a PDF file, simply do:
        //$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$settings['name'].'.xlsx"');
        header('Cache-Control: max-age=0');

        $Writer->save('php://output');
        exit;
    }
}