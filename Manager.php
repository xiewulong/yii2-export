<?php
/*!
 * yii2 extension - export
 * xiewulong <xiewulong@vip.qq.com>
 * https://github.com/xiewulong/yii2-export
 * https://raw.githubusercontent.com/xiewulong/yii2-export/master/LICENSE
 * create: 2015/8/5
 * update: 2015/8/5
 * version: 0.0.1
 */

namespace yii\export;

use Yii;
use PHPExcel;
use PHPExcel_IOFactory;

class Manager{

	//作者
	public $author;

	//生成对象
	private $creator;

	//输出对象
	private $factory;

	/**
	 * 生成文件
	 * @method create
	 * @since 0.0.1
	 * @param {array} $data 数据
	 * @param {string} $name 文件名
	 * @param {string} $type 文件类型
	 * @return {none}
	 * @example \Yii::$app->export->create($data, $name, $type);
	 */
	public function create($data, $name, $type = 'excel'){

		return 'yii2-export';

	}

}
