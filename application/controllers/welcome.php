<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

class Welcome extends CI_Controller 
{

	private $alignment_general;
	private $style_sidebar;
	private $style_head;
	private $style_person;
	private $style_dni;
	private $style_foot;
	private $line_firm;
	private $fotocheck_border;

	function __construct()
	{
		parent::__construct();

		$this->load->library('PHPExcel');
		$this->load->model('fotocheck_model');

		ini_set("memory_limit", "1024M");

	}
	
	public function index()
	{
		$this->load->view('welcome_message');
	}

	function set_styles()
	{
		$this->alignment_general = array(
			'alignment' => array(
				// 'wrap' => true,
				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
				'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
			),
		);

		$this->style_head = array(
			'alignment' => array(
				'vertical' => PHPExcel_Style_Alignment::VERTICAL_BOTTOM	
			),
			'font' => array(
				'name' => 'Arial',
				'size' => 9
			)
		);

		$this->style_person = array(
			'font' => array(
				'size' => 20,
				'italic' => true,
				'bold' => true
			)
		);

		$this->style_dni = array(
			'font' => array(
				'italic' => true,
				'bold' => true

			)
		);

		$this->style_foot = array(
			'font' => array(
				'size' => 9,
				'italic' => true
			)
		);

		$this->fotocheck_border = array(
			'borders' => array(
				'top' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN
				),
				'right' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN
				),
				'left' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN
				),
				'bottom' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN
				)
			)
		);

		$this->line_firm = array(
			'borders' => array(
				'bottom' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN
				)
			)
		);

	}

	function cell_value_with_merge($cell, $content, $merge)
	{
		$this->sheet->setCellValue($cell,$content);
		$this->sheet->mergeCells($merge);
	}

	function sidebar( $back_color )
	{
		$this->style_sidebar = array(
			'alignment' => array(
				'rotation' => 90
			),
			'font' => array(
				'size' => 34,
				'color' => array('rgb' => 'FFFFFF')
			),
			'borders' => array(
				'allborders' => array(
					'style' => PHPExcel_Style_Border::BORDER_THIN
				)
			),
			'fill' => array(
				'type' => PHPExcel_Style_Fill::FILL_SOLID,
				'color' => array('rgb' => $back_color )
			)
		);

		return $this->style_sidebar;
	}

	public function generate( $cod_sede_operativa )
	{
		////////////////////////////////
		//Colores y Estilos
		////////////////////////////////
		$this->set_styles();


		////////////////////////////////
		// Sheet 1
		////////////////////////////////
		$nro_sheet = 0;
		$sql = "SELECT p.dni, p.ape_paterno, p.ape_materno, p.nombres, c.id_cargo, c.cargo, c.cargo_res FROM PERSONAL p INNER JOIN CARGO c ON p.id_cargo = c.id_cargo WHERE c.id_cargo = 1 and p.cod_sede_operativa = $cod_sede_operativa ORDER BY p.ape_paterno ASC";
		$back_color = '366092';
		$name_sheet = 'APLICADOR';

		$valores = array( 'nro_sheet' => $nro_sheet, 'sql' => $sql, 'back_color' => $back_color, 'name_sheet' => $name_sheet );

		$this->sheet_base( $valores );


		////////////////////////////////
		// Sheet 2
		////////////////////////////////
		$nro_sheet = 1;
		$sql = "SELECT p.dni, p.ape_paterno, p.ape_materno, p.nombres, c.id_cargo, c.cargo, c.cargo_res FROM PERSONAL p INNER JOIN CARGO c ON p.id_cargo = c.id_cargo WHERE c.id_cargo = 2 and p.cod_sede_operativa = $cod_sede_operativa ORDER BY p.ape_paterno ASC";
		$back_color = '948A54';
		$name_sheet = 'ORIENTADOR';

		$valores = array( 'nro_sheet' => $nro_sheet, 'sql' => $sql, 'back_color' => $back_color, 'name_sheet' => $name_sheet );

		$this->sheet_base( $valores );


		////////////////////////////////
		// Sheet 3
		////////////////////////////////
		$nro_sheet = 2;
		$sql = "SELECT p.dni, p.ape_paterno, p.ape_materno, p.nombres, c.id_cargo, c.cargo, c.cargo_res FROM PERSONAL p INNER JOIN CARGO c ON p.id_cargo = c.id_cargo WHERE c.id_cargo = 3 and p.cod_sede_operativa = $cod_sede_operativa ORDER BY p.ape_paterno ASC";
		$back_color = '31869B';
		$name_sheet = 'ACL';

		$valores = array( 'nro_sheet' => $nro_sheet, 'sql' => $sql, 'back_color' => $back_color, 'name_sheet' => $name_sheet );

		$this->sheet_base( $valores );


		////////////////////////////////
		// Sheet 4
		////////////////////////////////
		$nro_sheet = 3;
		$sql = "SELECT p.dni, p.ape_paterno, p.ape_materno, p.nombres, c.id_cargo, c.cargo, c.cargo_res FROM PERSONAL p INNER JOIN CARGO c ON p.id_cargo = c.id_cargo WHERE c.id_cargo = 4 and p.cod_sede_operativa = $cod_sede_operativa ORDER BY p.ape_paterno ASC";
		$back_color = '538DD5';
		$name_sheet = 'INFORMATICO';

		$valores = array( 'nro_sheet' => $nro_sheet, 'sql' => $sql, 'back_color' => $back_color, 'name_sheet' => $name_sheet );

		$this->sheet_base( $valores );


		////////////////////////////////
		// Sheet 5
		////////////////////////////////
		$nro_sheet = 4;
		$sql = "SELECT p.dni, p.ape_paterno, p.ape_materno, p.nombres, c.id_cargo, c.cargo, c.cargo_res FROM PERSONAL p INNER JOIN CARGO c ON p.id_cargo = c.id_cargo WHERE c.id_cargo = 5 and p.cod_sede_operativa = $cod_sede_operativa ORDER BY p.ape_paterno ASC";
		$back_color = '60497A';
		$name_sheet = 'OPERADOR';

		$valores = array( 'nro_sheet' => $nro_sheet, 'sql' => $sql, 'back_color' => $back_color, 'name_sheet' => $name_sheet );

		$this->sheet_base( $valores );


		
		$this->phpexcel->getProperties()
		->setTitle("INEI - FOTOCHECK")
		->setDescription("Fotocheck");


		header("Content-Type: application/vnd.ms-excel");
		$nombreArchivo = 'FOTOCHECK_'.date('Y-m-d');
		header("Content-Disposition: attachment; filename=\"$nombreArchivo.xls\""); 
		header("Cache-Control: max-age=0");
		
		// Genera Excel
		$writer = PHPExcel_IOFactory::createWriter($this->phpexcel, "Excel5");

		$writer->save('php://output');
		exit;
	}

	public function sheet_base( $variable_array )
	{

		if ( $variable_array['nro_sheet'] == 0 )
		{
			// pestaña
			$this->sheet = $this->phpexcel->getActiveSheet(0);
		}
		else
		{
			$this->sheet = $this->phpexcel->createSheet( $variable_array['nro_sheet'] );
		}
		

		////////////////////////////////
		// Formato de la hoja ( Set Orientation, size and scaling )
		////////////////////////////////
		$this->sheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);// horizontal
		$this->sheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
		$this->sheet->getDefaultStyle()->getFont()->setName('Calibri');
		$this->sheet->getDefaultStyle()->applyFromArray($this->alignment_general);
		$this->sheet->getSheetView()->setZoomScale(100);
		$this->sheet->getDefaultColumnDimension()->setWidth(9.1); //default size column
		$this->sheet->getDefaultRowDimension()->setRowHeight(13.6);


		////////////////////////////////
		// Cuerpo
		////////////////////////////////

		$indice = 2; //fila inicial
		$contador = 0;

		$sql = $variable_array['sql'];
		$query = $this->convert_utf8->convert_result( $this->fotocheck_model->only_query( $sql ) );

		foreach ($query as $key => $row) 
		{
			if ( $contador % 2 != 0 )
			{
				//impar
				$column_sidebar = 'H'; 
				$column_start = 'I';
				$column_end = 'L';
				$column_logo = 'J';
				$column_firm_start = 'J';
				$column_firm_end = 'K';
			}
			else
			{
				//par
				$column_sidebar = 'B';
				$column_start = 'C';
				$column_end = 'F';
				$column_logo = 'D';
				$column_firm_start = 'D';
				$column_firm_end = 'E';
			}

			////////////////////////////////
			// SideBar
			////////////////////////////////
			$this->sheet->getColumnDimension($column_sidebar)->setWidth(8);
			$this->cell_value_with_merge( $column_sidebar.$indice, $row['cargo_res'], $column_sidebar.$indice.':'.$column_sidebar.($indice + 24) );
			$this->sheet->getStyle( $column_sidebar.$indice.':'.$column_sidebar.($indice + 24) )->applyFromArray( $this->sidebar( $variable_array['back_color'] ) );

			

			////////////////////////////////
			// Logo
			////////////////////////////////
			$objDrawing = new PHPExcel_Worksheet_Drawing();
			$objDrawing->setWorksheet($this->sheet);
			$objDrawing->setName("inei");
			$objDrawing->setDescription("Inei");
			$objDrawing->setPath("assets/img/inei.jpeg");
			$objDrawing->setCoordinates($column_logo.$indice);
			$objDrawing->setWidth(128);
			$objDrawing->setOffsetX(0);
			$objDrawing->setOffsetY(2);


			////////////////////////////////
			// Title
			////////////////////////////////
			$title_line1 = $indice + 5;
			$title_line2 = $indice + 6;
			$title_line3 = $indice + 7;
			$title_line4 = $indice + 8;

			$this->cell_value_with_merge( $column_start.$title_line1, 'EVALUACIÓN DEL CONCURSO PARA EL ', $column_start.$title_line1.':'.$column_end.$title_line1 );

			$this->cell_value_with_merge( $column_start.$title_line2, 'ACCESO A CARGOS DE DIRECTOR O SUB ', $column_start.$title_line2.':'.$column_end.$title_line2 );

			$this->cell_value_with_merge( $column_start.$title_line3, 'DIRECTOR DE INSTITUCIONES ', $column_start.$title_line3.':'.$column_end.$title_line3 );

			$this->cell_value_with_merge( $column_start.$title_line4, 'EDUCATIVAS PÚBLICAS', $column_start.$title_line4.':'.$column_end.$title_line4 );

			$this->sheet->getStyle( $column_start.$title_line1.':'.$column_end.$title_line4 )->applyFromArray( $this->style_head );

			////////////////////////////////
			// Person
			////////////////////////////////
			$names = $indice + 11;
			$last_name_1 = $indice + 12;
			$last_name_2 = $indice + 13;

			// $text_surname = trim( $row['ape_paterno'] ). ' ' . trim( $row['ape_materno'] );

			$this->cell_value_with_merge( $column_start.$names, trim($row['nombres']), $column_start.$names.':'.$column_end.$names );
			$this->cell_value_with_merge( $column_start.$last_name_1, trim( $row['ape_paterno'] ), $column_start.$last_name_1.':'.$column_end.$last_name_1 );
			$this->cell_value_with_merge( $column_start.$last_name_2, trim( $row['ape_materno'] ), $column_start.$last_name_2.':'.$column_end.$last_name_2 );

			$this->sheet->getStyle( $column_start.$names.':'.$column_end.$last_name_2 )->applyFromArray( $this->style_person );
			$this->sheet->getRowDimension($names)->setRowHeight(29);
			$this->sheet->getRowDimension($last_name_1)->setRowHeight(29);
			$this->sheet->getRowDimension($last_name_2)->setRowHeight(29);

			// DNI
			$dni = $indice + 14;
			$this->cell_value_with_merge( $column_start.$dni, 'D.N.I. N° '.$row['dni'], $column_start.$dni.':'.$column_end.$dni );

			$this->sheet->getStyle( $column_start.$dni.':'.$column_end.$dni )->applyFromArray( $this->style_dni );

			// Validez
			$validez = $indice + 15;
			$this->cell_value_with_merge( $column_start.$validez, 'VÁLIDO: 14 DE DICIEMBRE DE 2014', $column_start.$validez.':'.$column_end.$validez );


			////////////////////////////////
			// foot
			////////////////////////////////
			$firm = $indice + 20;
			$this->sheet->getStyle( $column_firm_start.$firm.':'.$column_firm_end.$firm )->applyFromArray( $this->line_firm );

			$foot_line1 = $indice + 21;
			$foot_line2 = $indice + 22;
			$foot_line3 = $indice + 23;

			$this->sheet->getRowDimension($foot_line1)->setRowHeight(10);
			$this->sheet->getRowDimension($foot_line2)->setRowHeight(10);
			$this->sheet->getRowDimension($foot_line3)->setRowHeight(10);

			$this->cell_value_with_merge( $column_start.$foot_line1, 'Director', $column_start.$foot_line1.':'.$column_end.$foot_line1 );
			$this->cell_value_with_merge( $column_start.$foot_line2, 'Oficina Departamental de Estadística e ', $column_start.$foot_line2.':'.$column_end.$foot_line2 );
			$this->cell_value_with_merge( $column_start.$foot_line3, 'Informática', $column_start.$foot_line3.':'.$column_end.$foot_line3 );


			$this->sheet->getStyle( $column_start.$foot_line1.':'.$column_end.$foot_line3 )->applyFromArray( $this->style_foot );

			$end_foot = $foot_line3 + 1;

			$this->sheet->getStyle( $column_start.$indice.':'.$column_end.$end_foot )->applyFromArray( $this->fotocheck_border );

			////////////////////////////////
			// Fondo
			////////////////////////////////
			$objDrawing = new PHPExcel_Worksheet_Drawing();
			$objDrawing->setWorksheet($this->sheet);
			$objDrawing->setName("inei");
			$objDrawing->setDescription("Inei");
			$objDrawing->setPath("assets/img/FONDO_INEI3.png");
			$objDrawing->setCoordinates($column_start.$indice);
			$objDrawing->setResizeProportional(false);
			$objDrawing->setWidth(255);
			$objDrawing->setHeight(520);
			$objDrawing->setOffsetX(0);
			$objDrawing->setOffsetY(0);

			if ( $contador > 0 && $contador % 2 != 0 )
			{
				$indice = $indice + 34;
			}

			$contador++;

		}

		////////////////////////////////
		// SALIDA EXCEL ( Propiedades del archivo excel )
		////////////////////////////////
		$this->sheet->setTitle( $variable_array['name_sheet'] );
		
	}

}

/* End of file welcome.php */
/* Location: ./application/controllers/welcome.php */