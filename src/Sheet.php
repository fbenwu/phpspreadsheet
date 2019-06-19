<?php
/**
 * Created by PhpStorm.
 * User: fan
 * Date: 2019/6/18
 * Time: 11:21
 * 因为每次处理表格文档都要花时间去找文档
 * 干脆花时间自己整理一个常用的phpspreadsheet读取和导出操作的类
 */
namespace Fan1992\Phpspreadsheet;

class Sheet
{
    const TYPE_XLS = 'xls'; // 导出类型 xls
    const TYPE_XLSX = 'xlsx'; // 导出类型 xlsx
    const TYPE_CSV = 'csv'; // 导出类型 csv

    public $filePath;
    public $readFirstLine = false;//是否读取首行
    public $down = true; //是否直接下载，false则保存文件在服务器上
    public $readDataOnly = true; // 不区分日期
    public $sheetNames = null; // 需要读取的表格名

    private $spreadsheets = null; // 文档对象
    private $activesheet = null; //当前sheet

    private $reader = null; // Reader
    private $writer = null; // Writer

    public function __construct()
    {
        set_time_limit(0); // 设置超时时间
    }

    /***
     * @param array $data
     * @param array $header
     * @param string $fileName
     * @param null $type
     * @param null $sheetNmae
     * @param array $width
     * @throws Exception
     * 导出
     */
    public function export($data = [], $header = [], $fileName = 'example', $type = null, $sheetNmae = null, $width = [])
    {
        if ($type && !in_array($type, [self::TYPE_CSV, self::TYPE_XLS, self::TYPE_XLSX])) {
            throw new \Exception('不支持的导出格式');
        }
        $this->spreadsheets = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $type               = $type ?: self::TYPE_XLSX;
        $this->setWriter($type);

        try {
            $this->activesheet = $this->spreadsheets->getActiveSheet();
            $this->writeSheet($data, $header, $sheetNmae, $width);
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new \Exception($e->getMessage());
        }
        $this->output($fileName, $type);
    }

    /***
     * @param array $data 包含每个sheet各项具体数据
     * @param string $fileName
     * @param null $type
     * @throws Exception
     * 多sheet导出
     */
    public function mutiSheetExport($data, $fileName = 'example', $type = null)
    {
        if ($type && !in_array($type, [self::TYPE_CSV, self::TYPE_XLS, self::TYPE_XLSX])) {
            throw new \Exception('不支持的导出格式');
        }
        $this->spreadsheets = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $type               = $type ?: self::TYPE_XLSX;
        $this->setWriter($type);

        try {
            if ($data) {
                $isFirstSheet = true;
                foreach ($data as $k => $sheetData) {
                    if ($isFirstSheet == true) {
                        $this->activesheet = $this->spreadsheets->getActiveSheet();
                    } else {
                        $this->activesheet = $this->spreadsheets->createSheet();
                    }
                    $temData   = is_array($sheetData) && isset($sheetData['data']) && $sheetData['data'] ? $sheetData['data'] : [];
                    $temHeader = is_array($sheetData) && isset($sheetData['header']) && $sheetData['header'] ? $sheetData['header'] : [];
                    $temWidth  = is_array($sheetData) && isset($sheetData['width']) && $sheetData['width'] ? $sheetData['width'] : [];
                    $temName   = is_array($sheetData) && isset($sheetData['sheetName']) && $sheetData['sheetName'] ? $sheetData['sheetName'] : 'Sheet' . strval($k + 1);
                    $this->writeSheet($temData, $temHeader, $temName, $temWidth);
                    $isFirstSheet = false;
                }
            }
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            throw new \Exception($e->getMessage());
        }

        $this->output($fileName, $type);
    }

    /***
     * @param $fileName
     * @param $ext
     * @param null $savePath
     * 输出
     * 本地保存或浏览器下载
     */
    private function output($fileName, $ext, $savePath = null)
    {
        if ($this->down) {
            $this->setOutputHeader($ext, $fileName);
            //清除缓存
            ob_clean();
            $this->writer->save('php://output');
        } else {
            $file = $savePath . '/' . $fileName . '.' . $ext;
            $this->writer->save($file);
        }
    }

    /***
     * @param $ext
     * @param $fileName
     * 设置浏览器输出的响应头
     */
    private function setOutputHeader($ext, $fileName)
    {
        header("Content-Type: application/vnd.ms-excel; charset=UTF8");
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Content-Type: application/force-download");
        header("Content-Type: application/octet-stream");
        header("Content-Type: application/download");
        header("Content-Disposition: attachment;filename=" . $fileName . '.' . $ext);
        header("Content-Transfer-Encoding: binary ");
    }

    /**
     * @param array $data
     * @param array $header
     * @param null $sheetName
     * 将数据写入当前 active sheet
     */
    private function writeSheet($data = [], $header = [], $sheetName = null, $whith = [])
    {
        // sheet name
        $sheetName = $sheetName ?: 'Sheet1';
        $this->activesheet->setTitle($sheetName);

        // header
        if ($header && is_array($header)) {
            $this->setActiveSheetHeader($header);
        }

        // data
        if ($data && is_array($data)) {
            $this->setActiveSheetData($data, $header ? true : false);
        }

        $this->setActiveSheetColumnWidth($whith);

    }

    /***
     * @param $width
     * 列宽
     */
    private function setActiveSheetColumnWidth($width)
    {
        $col     = 1;
        $maxCols = $this->activesheet->getHighestColumn();
        for ($i = 1; $i <= ord($maxCols) - ord('A') + 1; $i++) {
            if (is_array($width) && isset($width[$i - 1]) && intval($width[$i - 1])) {
                $this->activesheet->getColumnDimensionByColumn($col)->setWidth($width[$i - 1]);
            } else {
                $this->activesheet->getColumnDimensionByColumn($col)->setAutoSize(true); // 自适应宽度
            }
            $col++;
        }
    }

    /***
     * @param $data
     * @param bool $hasHeader
     * 写入active sheet具体数据
     */
    private function setActiveSheetData($data, $hasHeader = true)
    {
        $currentRow = $hasHeader ? 2 : 1;
        if (!$data) {
            return;
        }
        foreach ($data as $key => $rows) {
            $rowspan           = 1;//当前数据所需要合并的最大单元格行数
            $rows_division_arr = [];//单元格合并数
            if ($rows) {//找出当前记录要合并的最大单元格行数
                foreach ($rows as $k => $v) {
                    $tem                   = 0;
                    $tem_rows_division_arr = [];
                    if (is_array($v)) {
                        foreach ($v as $vv) {
                            if (is_array($vv)) {
                                $tem                     += count($vv);
                                $tem_rows_division_arr[] = count($vv);
                            } else {
                                $tem++;
                            }
                        }
                    }
                    if ($tem > $rowspan) {
                        $rowspan           = $tem;
                        $rows_division_arr = $tem_rows_division_arr;
                    }
                }
            }

            foreach ($rows as $index => $value) {
                $j = intval($index + 1);
                if (!is_array($value)) {
                    if ($rowspan > 1) {
                        $this->activesheet->mergeCellsByColumnAndRow($j, $currentRow, $j, intval($currentRow + $rowspan - 1));
                    }
                    $this->activesheet->setCellValueByColumnAndRow($j, $currentRow, $value);
                    $this->setCellCenter($j, $currentRow);
                } else {
                    $tem_column = $currentRow;
                    foreach ($value as $k => $v) {
                        if (!is_array($v)) {
                            if ($rows_division_arr) {
                                if ($rows_division_arr[$k] > 1) {
                                    $this->activesheet->mergeCellsByColumnAndRow($j, $tem_column, $j, intval($tem_column + $rows_division_arr[$k] - 1));
                                }
                                $this->activesheet->setCellValueByColumnAndRow($j, $tem_column, $v);
                                $this->setCellCenter($j, $tem_column);
                                $tem_column += $rows_division_arr[$k];
                            } else {
                                $this->activesheet->setCellValueByColumnAndRow($j, $tem_column, $v);
                                $this->setCellCenter($j, $tem_column);
                                $tem_column++;
                            }
                        } else {
                            foreach ($v as $vv) {
                                $this->activesheet->setCellValueByColumnAndRow($j, $tem_column, $vv);
                                $this->setCellCenter($j, $tem_column);
                                $tem_column++;
                            }
                        }
                    }
                }
            }
            $currentRow += $rowspan;
        }
    }

    /**
     * @param $col
     * @param $row
     * 设置单元格居中
     */
    private function setCellCenter($col, $row)
    {
        $styleArray = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ];
        $this->activesheet->getStyleByColumnAndRow($col, $row)->applyFromArray($styleArray);
    }

    /***
     * @param $header
     * 设置active sheet 表头
     */
    private function setActiveSheetHeader($header)
    {
        foreach ($header as $k => $v) {
            $this->activesheet->setCellValueByColumnAndRow(intval($k + 1), 1, $v);
            $this->setCellCenter(intval($k + 1), 1);
        }
    }

    /***
     * @param $type
     * @throws Exception
     * 自动适配 Writer
     */
    private function setWriter($type)
    {
        switch ($type) {
            case self::TYPE_XLSX:
                $this->writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheets);
                break;
            case self::TYPE_XLS:
                $this->writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheets);
                break;
            case self::TYPE_CSV:
                $this->writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($this->spreadsheets);
                break;
            default:
                throw new \Exception('不支持的导出格式');
        }
    }

    /***
     * @param $file 文件名称，包含路径
     * @return array
     * @throws Exception
     * 读取表格文件，包括 xls，xlsx，csv，其他没有测试
     * 支持读一个文件里面的多个sheet
     */
    public function read($file)
    {
        if (!$file) {
            $file = $this->filePath;
        }
        if (!file_exists($file)) {
            throw new \Exception('文件不存在');
        }
        $arr = pathinfo($file);
        $ext = $arr['extension'];
        $this->setReader($ext);
        $this->reader->setReadDataOnly($this->readDataOnly);
        if (!$this->sheetNames) {
            $this->reader->setLoadAllSheets();
        } else {
            $this->reader->setLoadSheetsOnly($this->sheetNames);
        }
        $this->spreadsheets = $this->reader->load($file);

        $data = [];
        if (!$this->sheetNames) {
            $this->activesheet = $this->spreadsheets->getActiveSheet();
            $data              = $this->getActiveSheetData();
        } else {
            $isSingleSheet = !(count($this->sheetNames) > 1);
            foreach ($this->sheetNames as $k => $v) {
                if (!$this->spreadsheets->getSheetByName($v)) {
                    throw new \Exception('工作表 ' . $v . ' 不存在');
                }
                $this->spreadsheets->setActiveSheetIndexByName($v);
                $this->activesheet = $this->spreadsheets->getActiveSheet();
                if ($isSingleSheet) {
                    $data = $this->getActiveSheetData();
                } else {
                    $data[$k] = $this->getActiveSheetData();
                }
            }
        }

        return $data;
    }

    /***
     * @return array
     * 获取当前活动sheet数据
     */
    private function getActiveSheetData()
    {
        $highestRow         = $this->activesheet->getHighestRow();  // 最大行数
        $highestColumn      = $this->activesheet->getHighestColumn(); // 最大列数
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

        $data = [];
        for ($row = 1; $row <= $highestRow; $row++) {
            $lineData = [];
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $lineData[] = $this->activesheet->getCellByColumnAndRow($col, $row)->getValue();
            }
            $data[] = $lineData;
        }
        return $data;
    }

    /***
     * @param $type
     * @return null|\PhpOffice\PhpSpreadsheet\Reader\Csv|\PhpOffice\PhpSpreadsheet\Reader\Ods|\PhpOffice\PhpSpreadsheet\Reader\Slk|\PhpOffice\PhpSpreadsheet\Reader\Xls|\PhpOffice\PhpSpreadsheet\Reader\Xlsx
     * @throws Exception
     * 根据文件后缀适配Reader
     */
    private function setReader($type)
    {
        switch (strtolower($type)) {
            case 'xls':
                $this->reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
                break;
            case 'xlsx':
                $this->reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                break;
            case 'csv':
                $this->reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
                break;
            case 'ods':
                $this->reader = new \PhpOffice\PhpSpreadsheet\Reader\Ods();
                break;
            case 'slk':
                $this->reader = new \PhpOffice\PhpSpreadsheet\Reader\Slk();
                break;
            case 'gnumeric':
                $this->reader = new \PhpOffice\PhpSpreadsheet\Reader\Gnumeric();
                break;
            default:
                throw new \Exception('不支持的格式');
        }
    }
}

?>