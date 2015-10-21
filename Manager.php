<?php
/*!
 * yii2 extension - export
 * xiewulong <xiewulong@vip.qq.com>
 * https://github.com/xiewulong/yii2-export
 * https://raw.githubusercontent.com/xiewulong/yii2-export/master/LICENSE
 * create: 2015/8/5
 * update: 2015/10/21
 * version: 0.0.1
 */

namespace yii\export;

use Yii;
use PHPExcel;
use PHPExcel_IOFactory;

class Manager{

	//作者
	public $author;

	//导出文件类型, 默认导出excel
	private $type = 'excel';

	//生成对象
	private $creator;

	//Excel列标识集合
	private $excelColumnNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

	/**
	 * 生成文件
	 * @method create
	 * @since 0.0.1
	 * @param {string} $name 文件名
	 * @param {array} $titles 标题
	 * @param {array} $datas 数据
	 * @param {function} $handler 数据处理方法
	 * @param {string} $sheet_title 标签页标题
	 * @param {number} $sheet_index 标签页索引
	 * @return {none}
	 * @example \Yii::$app->export->create($name, $titles, $datas, $handler, $sheet_title, $sheet_index);
	 */
	public function create($name, $titles, $datas, $handler, $sheet_title = null, $sheet_index = 0){
		$creator = $this->getCreator();
		$sheet = $creator->setActiveSheetIndex($sheet_index);

		$row = 1;
		foreach($titles as $index => $title){
			$sheet->setCellValue($this->getExcelColumnName($index) . $row, $title);
		}

		$row++;
		foreach($datas as $data){
			list($data) = $_data = call_user_func($handler, $data);
			$unmerge = isset($_data[1]) ? $_data[1] : null;
			if($merge = $unmerge && isset($data[$unmerge])){
				$_row = $row + count($data[$unmerge]) - 1;
			}
			$col = 0;
			foreach($data as $key => $value){
				if($merge && $key === $unmerge){
					$_col = $col;
					foreach($value as $index => $__data){
						foreach($__data as $_index => $_value){
							$sheet->setCellValue($this->getExcelColumnName($_col + $_index) . ($row + $index), $_value);
						}
					}
					$col += $_index;
				}else{
					$colName = $this->getExcelColumnName($col);
					$sheet->setCellValue("$colName$row", $value);
					if($merge && $_row > $row){
						$sheet->mergeCells("$colName$row:$colName$_row");
					}
				}
				$col++;
			}
			if($merge && $_row > $row){
				$row = $_row;
			}
			$row++;
		}

		if($sheet_title){
			$creator->getActiveSheet()->setTitle($sheet_title);
		}

		return $this->output($name);
	}

	/**
	 * 设置导出文件类型
	 * @method setFileType
	 * @since 0.0.1
	 * @param {number} $type 导出文件类型
	 * @return {none}
	 * @example $this->setFileType($type);
	 */
	public function setFileType($type){
		$this->type = $type;
	}

	/**
	 * 输出文件
	 * @method output
	 * @since 0.0.1
	 * @param {string} $filename 文件名称
	 * @return {none}
	 * @example $this->output($filename);
	 */
	private function output($filename){
		switch($this->type){
			case 'excel':
				header('Content-Type: application/vnd.ms-excel');
				header("Content-Disposition: attachment; filename=$filename.xls");
				header('Cache-Control: max-age=0');
				// If you're serving to IE 9, then the following may be needed
				header('Cache-Control: max-age=1');
				// If you're serving to IE over SSL, then the following may be needed
				header ('Expires: ' . gmdate('D, d M Y H:i:s', 0) . ' GMT');	// Date in the past
				header ('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');	// always modified
				header ('Cache-Control: cache, must-revalidate');	// HTTP/1.1
				header ('Pragma: public');	// HTTP/1.0
				$factory = PHPExcel_IOFactory::createWriter($this->creator, 'Excel5');
				return $factory->save('php://output');
				break;
		}
	}

	/**
	 * 设置生成对象
	 * @method getCreator
	 * @since 0.0.1
	 * @return {none}
	 * @example $this->getCreator();
	 */
	private function getCreator(){
		if(!$this->creator){
			switch($this->type){
				case 'excel':
					$this->creator = new PHPExcel;
					$this->creator->getProperties()
						->setCreator($this->author)
						->setLastModifiedBy($this->author)
						->setTitle($this->author);
					break;
			}
		}

		return $this->creator;
	}

	/**
	 * 获取序号标识
	 * @method getExcelColumnName
	 * @since 0.0.1
	 * @param {number} $index 序号
	 * @return {string}
	 * @example $this->getExcelColumnName($index);
	 */
	private function getExcelColumnName($index){
		$len = count($this->excelColumnNames);
		$col = [];

		while($index >= 0){
			array_unshift($col, $this->excelColumnNames[$index % $len]);
			$index = floor($index / $len) - 1;
		}

		return implode('', $col);
	}

}
