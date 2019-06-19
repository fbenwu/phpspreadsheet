<?php
/**
 * Created by PhpStorm.
 * User: wufan
 * Date: 2019/6/18
 * Time: 11:53
 */

//require_once '../vendor/autoload.php';

//use Fan1992\Phpspreadsheet\Sheet;

require_once './../src/Sheet.php';
$sheet = new Sheet();

// 多sheet读取
//$sheet->sheetNames = ['data_a'=>'表1','data_b'=>'表2','data_c'=>'Sheet3'];
//$data = $sheet->read('./test.xlsx');
//print_r($data);exit;

// 导出
//$header = ['提提1', 'title2', '标题3', '测试测试'];
//$data = [
//    ['12', ['阿斯顿发','asdf','2019-06-19'], '是的', '沙箱'],
//    ['撒发顺丰的', ['1','23','撒旦法师'], '2019-06-19', 'asdasdfas']
//];
//$width = [30,0,40,60];
//$sheet->export($data, $header,'test'.time(),null, 'mysheet', $width);

// 多sheet 导出
$sheet1Data = [
    'data'      => [
        ['12', ['阿斯顿发', 'asdf', '2019-06-19'], '是的', '沙箱'],
        ['asd', ['阿斯顿发', '22', '2019-06-19'], '是', '沙asdf箱'],
    ],
    'header'    => ['header1', '标题2', '333', '超级长超级长超级长超级长超级长超级长超级长的标题'],
    'sheetName' => '1211sheet1',
    'width'     => [5, 0, 5]
];
$sheet2Data = [
    'data'      => [
        ['12', ['阿斯顿发', 'asdf', '2019-06-19'], '是的', '沙箱'],
        ['asd', ['阿斯顿发', '22', '2019-06-19'], '是', '沙asdf箱'],
    ],
    'header'    => ['header1', '标题2', '333', '超级长超222222级长超级长超级长超级长超级长超级长的标题'],
    'sheetName' => '',
    'width'     => []
];
$sheet3Data = [
    'data'      => [
        ['12', ['阿斯顿发', 'asdf', '2019-06-19'], '是的', '沙333箱'],
        ['asd', ['阿斯顿发', '22', '2019-06-19'], '是3333', '沙asdf箱'],
    ],
    'header'    => ['header1', '标题2', '333', '超级长超级长超3333级长超级长超级长超级长超级长的标题'],
    'sheetName' => '导出sheet3',
    'width'     => []
];
$data       = [$sheet1Data, $sheet2Data, $sheet3Data];
$sheet->mutiSheetExport($data, 'muti' . time());