<?php
/**
 * Created by PhpStorm.
 * User: fan
 * Date: 2019/6/20
 * Time: 10:12
 */

namespace Fan1992\Phpspreadsheet;

class Csv
{
    public $readFirstLine = false;//是否读取首行
    public $down = true; //是否直接下载，false则保存文件在服务器上

    public function __construct()
    {
        ini_set('max_execution_time', 0); //设置程序的执行时间,0为无上限
        ini_set('memory_limit', '1024M'); //设置程序运行的内存
    }

    /***
     * @param $file 带路径的文件地址
     * @return array
     * 读取csv文件
     */
    public function read($file)
    {
        setlocale(LC_ALL, 'zh_CN');//linux系统下生效
        $data = [];//返回的文件数据行
        if (!is_file($file) && !file_exists($file)) {
            die('文件错误');
        }
        $cvsFile = fopen($file, 'r'); //开始读取csv文件数据
        $i        = 0;//记录cvs的行
        while ($fileData = fgetcsv($cvsFile)) {
            $i++;
            if ($i == 1 && !$this->readFirstLine) {
                continue;//过滤表头
            }
            if ($fileData[0] != '') {
                foreach ($fileData as &$v){
                    $v = iconv('GBK', 'UTF-8', $v);
                }
                $data[$i] = $fileData;
            }

        }
        fclose($cvsFile);
        return $data;
    }

    /***
     * 导出或保存csv文件
     * @param array $data
     * @param array $header
     * @param string $name
     * 注意：如果是保存而不下载，$name 带上保存路径
     */
    public function export($data = [], $header = [], $name = 'example')
    {
        $name = iconv('UTF-8', 'GBK', $name);
        if ($this->down) {
            $this->down($data, $header, $name);
        } else {
            $this->save($data, $header, $name);
        }
    }

    /***
     * 保存文件
     * @param array $data
     * @param array $header
     * @param string $name
     */
    private function save($data = [], $header = [], $name = 'example')
    {
        $file = $name . '.csv';
        $fp   = fopen($file, 'w');
        $this->write($fp, $data, $header);
    }

    /**
     * 向句柄中写入内容
     * @param $fp
     * @param $data
     * @param $header
     */
    private function write($fp, $data, $header)
    {
        fwrite($fp, chr(0xEF) . chr(0xBB) . chr(0xBF));
        fputcsv($fp, $header);
        $index = 0;
        foreach ($data as $item) {
            if ($index == 1000) { //每次写入1000条数据清除内存
                $index = 0;
                ob_flush();//清除内存
                flush();
            }
            $index++;
            if ($item) {
                foreach ($item as &$v) {
                    $v = iconv('UTF-8', 'GBK', $v);
                }
            }
            fputcsv($fp, $item);
        }
        @ob_flush();
        flush();
        ob_end_clean();
        fclose($fp);
        die();
    }

    /**
     * 下载文件
     * @param array $data
     * @param array $header
     * @param $name
     */
    private function down($data = [], $header = [], $name)
    {
        $this->setHeader($name);
        $fp = fopen('php://output', 'w');
        $this->write($fp, $data, $header);
    }

    /***
     * 设置浏览器下载的响应头
     * @param $name
     */
    private function setHeader($name)
    {
        header('Content-Description: File Transfer');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename=' . $name . '.csv');
        header('Cache-Control: max-age=0');
    }
}