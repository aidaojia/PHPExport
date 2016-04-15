<?php
/**
 * Created by PhpStorm.
 * User: zuoluo
 * Date: 16/4/15
 * Time: 下午4:18
 */
namespace Aidaojia\PHPExport;

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
    public function __construct($handleData)
    {
        if ( ! $handleData ) {
            $this->handleDate = $handleData;
        }
        $this->excel = new PHPExcel();
    }

    /**
     * PHPExport getExcelByData.
     */
    public function getExcelByData()
    {
        if ( ! $this->handleDate ) {
            $this->excel;
        }

        foreach ($this->handleDate as $key => $val) {
            // active sheet index
            $this->excel->setActiveSheetIndex($this->sheetIndex);

            // get sheet by given data
            $this->_getSheetByData($this->excel, $key, $val);
        }
    }

    /**
     *  PHPExport _getSheetByData
     */
    private function _getSheetByData($excel, $sheetTitle, $sheetValue)
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
        $endColumn = getMergeColumn($length-1);
        $sheet->mergeCells("A1:{$endColumn}1");
        $sheet->getStyle("A1:{$endColumn}1")->applyFromArray($commonStyle);
        $sheet->getRowDimension('1')->setRowHeight(20);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // generate sheet content not include A1 cell
        $this->_getSheetContent($sheet, $sheetValue);

        return $excel;
    }

    private function _getSheetContent(& $sheet, $sheetValue)
    {
        if ( ! $sheetValue ) {
            return;
        }

        $this->columnIndex = 1;
        foreach ($sheetValue as $key => $val) {

            ++ $this->columnIndex;
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