<?php
/**
 * Created by PhpStorm.
 * User: zuoluo
 * Date: 16/4/15
 * Time: 下午4:18
 *
 * 基于phpexcel导出
 * 数据格式定义array:
 * [ sheetTitle => [ columnIndex1 => [columnValue1, columnValue2]],
 *                   columnIndex2 => [columnValue1, columnValue2]],
 *                   columnIndex3 => [columnValue1, columnValue2]]
 *                  ];
 *
 * [ 所有产品列表 => [  title => [产品Id, 产品编码],
 *                    1252 =>  [1252, SG1251],
 *                    1253 =>  [1253, SG1251]
 *                  ];
 *
 * 表格标题定义:
 * 可支持自定义传入,否则数据日期格式标题
 * 构造方法中传入
 * public function __construct($handleData, $excelName = '')
 *
 */
namespace Aidaojia\PHPExport;

use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Style_Border;
use PHPExcel_Style_Alignment;

class PHPExport
{
    /**
     *
     * @var PHPExcel
     */
    private $excel = null;

    /**
     *
     * @var array $handleDate
     */
    private $handleDate = [];

    /**
     *
     * @var string $excelName
     */
    private $excelName = '';

    /**
     *
     * @var int $sheetIndex
     */
    private $sheetIndex = 0;

    /**
     *
     * @var int $columnIndex
     */
    private $columnIndex = 1;

    /**
     * PHPExport constructor.
     *
     * @param $handleData
     */
    public function __construct($handleData, $excelName = '')
    {
        if ( $handleData ) {
            $this->handleDate = $handleData;
        }

        if ( $excelName ) {
            $this->excelName = $excelName;
        } else {
            $this->excelName = date('Y-m-d', time()) . '表格导出';
        }

        $this->excel = new PHPExcel();
    }

    /**
     * PHPExport getExcelByData.
     */
    public function getExcelByData()
    {
        if ( ! $this->handleDate ) {
            return $this->excel;
        }

        foreach ($this->handleDate as $key => $val) {
            // active sheet index
            $this->excel->setActiveSheetIndex($this->sheetIndex ++);

            // get sheet by given data
            $this->_getSheetByData($this->excel, $key, $val);

            $this->excel->createSheet();
        }

        if (!headers_sent()) {
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename=' . $this->excelName . '.xlsx');
            header('Cache-Control: max-age=0');
            $output = PHPExcel_IOFactory::createWriter($this->excel, 'Excel2007');
            $output->save('php://output');
            exit;
        }
    }

    /**
     *  PHPExport _getSheetByData
     */
    private function _getSheetByData(& $excel, $sheetTitle, $sheetValue)
    {
        // get common style
        $commonStyle = $this->_getCommonStyle();

        // set A1 cell data
        $sheet = $excel->getActiveSheet();
        $sheet->setTitle($sheetTitle);
        $sheet->setCellValue('A1', $sheetTitle);

        // set A1 cell param
        $sheet->getStyle('A1')->getFont()->setBold(true);
        $length = count($sheetValue['title']);
        $endColumn = $this->_getMergeColumn($length-1);
        $sheet->mergeCells("A1:{$endColumn}1");
        $sheet->getStyle("A1:{$endColumn}1")->applyFromArray($commonStyle);
        $sheet->getRowDimension('1')->setRowHeight(20);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // generate sheet content not include A1 cell
        $this->_getSheetContent($sheet, $sheetValue);

    }

    private function _getSheetContent(& $sheet, $sheetValue)
    {
        if ( ! $sheetValue ) {
            return;
        }

        $this->columnIndex = 1;
        foreach ($sheetValue as $key => $val) {

            $this->columnIndex ++;
            $columnLength = count($val);

            for($i = 0; $i < $columnLength; $i++) {
                $columnLetter = $this->_getMergeColumn($i);
                $sheet->setCellValue("{$columnLetter}{$this->columnIndex}", $val[$i]);
                $sheet->getStyle("{$columnLetter}{$this->columnIndex}")->applyFromArray($this->_getCommonStyle());
            }
        }
    }

    /**
     * @return array commonStyle
     */
    private function _getCommonStyle()
    {
        $commonStyle = array(
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN
                ),
            ),
        );

        return $commonStyle;
    }

    /**
     *
     * 根据map数组获取对应的字母
     */
    private function _getMergeColumn($length) {
        if ($length < 0 || $length > 25) {
            echo "导出数据列长度不合法，不能小于1或者大于26";
        }

        $columnMap = [
            0 => 'A',
            1 => 'B',
            2 => 'C',
            3 => 'D',
            4 => 'E',
            5 => 'F',
            6 => 'G',
            7 => 'H',
            8 => 'I',
            9 => 'J',
            10 => 'K',
            11 => 'L',
            12 => 'M',
            13 => 'N',
            14 => 'O',
            15 => 'P',
            16 => 'Q',
            17 => 'R',
            18 => 'S',
            19 => 'T',
            20 => 'U',
            21 => 'V',
            22 => 'W',
            23 => 'X',
            24 => 'Y',
            25 => 'Z',
        ];

        return $columnMap[$length];
    }


}

