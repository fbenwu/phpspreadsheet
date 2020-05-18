# phpspreadshhet
phpspreadshhet 读取，导出

# 安装
        composer require fan1992/phpspreadsheet
# 使用示例
        use Fan1992\Phpspreadsheet\Excel;
        $excel = new Excel();
        
# 读取文档（默认只读取第一个sheet）
        $data = $excel->read('./test.xlsx');
        print_r($data);
# 一次读取多个sheet
        $excel->sheetNames = ['data_a'=>'表1','data_b'=>'表2','data_c'=>'Sheet3'];
        $data = $excel->read('./test.xlsx');
        print_r($data);
# 一次读取所有sheet
        $excel = new Excel();
        $excel->autoReadAllSheets = true;
        $data = $excel->read('./test.xlsx');
        print_r($data);
        
# 导出
        $header = ['提提1', 'title2', '标题3', '测试测试']; //表头，即第一行
        $data = [//具体数据，支持多维数组（合并单元格）
            ['12', ['阿斯顿发','asdf','2019-06-19'], '是的', '沙箱'],
            ['撒发顺丰的', ['1','23','撒旦法师'], '2019-06-19', 'asdasdfas']
        ];
        $width = [30,0,40,60]; //列宽度，为0或没有则自动适应
        $excel->export($data, $header,'test'.time(),null, 'mysheet', $width);
# 导出多个sheet
        $excel1Data = [
            'data'      => [
                ['12', ['阿斯顿发', 'asdf', '2019-06-19'], '是的', '沙箱'],
                ['asd', ['阿斯顿发', '22', '2019-06-19'], '是', '沙asdf箱'],
            ],
            'header'    => ['header1', '标题2', '333', '超级长超级长超级长超级长超级长超级长超级长的标题'],
            'sheetName' => '1211sheet1',
            'width'     => [5, 0, 5]
        ];
        $excel2Data = [
            'data'      => [
                ['12', ['阿斯顿发', 'asdf', '2019-06-19'], '是的', '沙箱'],
                ['asd', ['阿斯顿发', '22', '2019-06-19'], '是', '沙asdf箱'],
            ],
            'header'    => ['header1', '标题2', '333', '超级长超222222级长超级长超级长超级长超级长超级长的标题'],
            'sheetName' => '',
            'width'     => []
        ];
        $excel3Data = [
            'data'      => [
                ['12', ['阿斯顿发', 'asdf', '2019-06-19'], '是的', '沙333箱'],
                ['asd', ['阿斯顿发', '22', '2019-06-19'], '是3333', '沙asdf箱'],
            ],
            'header'    => ['header1', '标题2', '333', '超级长超级长超3333级长超级长超级长超级长超级长的标题'],
            'sheetName' => '导出sheet3',
            'width'     => []
        ];
        $data       = [$excel1Data, $excel2Data, $excel3Data];
        $excel->mutiSheetExport($data, 'muti' . time(), Sheet::TYPE_XLSX);
        
# 其它
        $readFirstLine = false;//是否读取首行
        $down = true; //是否直接下载，false则保存文件在服务器上
        
        支持导出格式：xls，xlsx
        
