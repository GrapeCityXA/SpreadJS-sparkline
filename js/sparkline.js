
var spreadNS = GC.Spread.Sheets;

window.onload = function () {
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
    //初始化迷你图的方法
    initSpread(spread);
};

//初始化迷你图的方法
function initSpread(spread) {
    //初始化表格
    var sheet = spread.getSheet(0);
    sheet.suspendPaint();
    sheet.options.allowCellOverflow = true;
    //设置数据
    var data = [1,-2,-1,6,4,-4,3,8];
    var dateAxis = [new Date(2011, 0, 5),new Date(2011, 0, 1),new Date(2011, 1, 11),new Date(2011, 2, 1),
            new Date(2011, 1, 1),new Date(2011, 1, 3),new Date(2011, 2, 6),new Date(2011, 1, 19)];
    sheet.setValue(0, 0, "Series 1");
    sheet.setValue(0, 1, "Series 2");
    //数据循环写入
    for(let i=0;i<8;i++)
    {
        sheet.setValue(i+1, 0,data[i]);
        sheet.getCell(i+1, 1).value(dateAxis[i]).formatter("yyyy-mm-dd");
    }
    //设置宽高
    sheet.setColumnWidth(1,100);
    sheet.setValue(11, 0, "*Data Range is A2-A9");
    sheet.setValue(12, 0, "*Date axis range is B2-B9");

    var dataRange = new spreadNS.Range(1, 0, 8, 1);
    var dateAxisRange = new spreadNS.Range(1, 1, 8, 1);

    //设置迷你图不包含日期坐标轴
    sheet.getCell(0, 5).text("Sparkline without dateAxis:");

    sheet.getCell(1, 5).text("(1) Line");
    sheet.getCell(1, 8).text("(2) Column");
    sheet.getCell(1, 11).text("(3) Winloss");

    //设置迷你图包含日期坐标轴
    sheet.getCell(7, 5).text("Sparkline with dateAxis:");

    sheet.getCell(8, 5).text("(1) Line");
    sheet.getCell(8, 8).text("(2) Column");
    sheet.getCell(8, 11).text("(3) Winloss");

    //迷你图优化设置
    var setting = new spreadNS.Sparklines.SparklineSetting();
    setting.options.showMarkers = true;
    setting.options.lineWeight = 3;
    setting.options.displayXAxis = true;
    setting.options.showFirst = true;
    setting.options.showLast = true;
    setting.options.showLow = true;
    setting.options.showHigh = true;
    setting.options.showNegative = true;
    //行
    sheet.addSpan(2, 5, 4, 3);
    //setSparkline方法用来设置单元格
    sheet.setSparkline(2, 5, dataRange
            , spreadNS.Sparklines.DataOrientation.vertical
            , spreadNS.Sparklines.SparklineType.line
            , setting
    );

    sheet.addSpan(9, 5, 4, 3);
    sheet.setSparkline(9, 5, dataRange
        , GC.Spread.Sheets.Sparklines.DataOrientation.vertical
        , GC.Spread.Sheets.Sparklines.SparklineType.line
            , setting
        , dateAxisRange
        , GC.Spread.Sheets.Sparklines.DataOrientation.vertical
    );

    //列
    sheet.addSpan(2, 8, 4, 3);
    sheet.setSparkline(2, 8, dataRange
            , spreadNS.Sparklines.DataOrientation.vertical
            , spreadNS.Sparklines.SparklineType.column
            , setting
    );



    sheet.addSpan(9, 8, 4, 3);
    sheet.setSparkline(9, 8, dataRange
        , GC.Spread.Sheets.Sparklines.DataOrientation.vertical
        , GC.Spread.Sheets.Sparklines.SparklineType.column
            , setting
        , dateAxisRange
        , GC.Spread.Sheets.Sparklines.DataOrientation.vertical
    );

    //winloss
    sheet.addSpan(2, 11, 4, 3);
    sheet.setSparkline(2, 11, dataRange
            , spreadNS.Sparklines.DataOrientation.vertical
            , spreadNS.Sparklines.SparklineType.winloss
            , setting
    );

    sheet.addSpan(9, 11, 4, 3);
    sheet.setSparkline(9, 11, dataRange
        , GC.Spread.Sheets.Sparklines.DataOrientation.vertical
        , GC.Spread.Sheets.Sparklines.SparklineType.winloss
            , setting
        , dateAxisRange
        , GC.Spread.Sheets.Sparklines.DataOrientation.vertical
    );

    sheet.bind(spreadNS.Events.SelectionChanged, selectionChangedCallback);

    sheet.resumePaint();

    //选择已更改的方法
    function selectionChangedCallback() {
        var sheet = spread.getActiveSheet();
        var sparkline = sheet.getSparkline(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
        //判断更新迷你图还是新增迷你图
        if (sparkline) {
            updateSetting(sparkline);
        } else {
            initSetting();
        }
    }

    //更新迷你图
    function updateSetting(sparkline) {
        var type = sparkline.sparklineType(), orientation = sparkline.dataOrientation(),
            row = sparkline.row, column = sparkline.column;
        _getElementById("line_position").value = row + "," + column;

        var line_type = _getElementById("line_type");
        _selectOption(line_type, type + "");

        var line_orientation = _getElementById("line_orientation");
        _selectOption(line_orientation, orientation + "");
    }

    //新增迷你图
    function initSetting() {
        _getElementById("line_position").value = '';

        var line_type = _getElementById("line_type");
        _selectOption(line_type, '0');

        var line_orientation = _getElementById("line_orientation");
        _selectOption(line_orientation, '0');
    }

    //获取实际单元格范围的方法
    function getActualCellRange(cellRange, rowCount, columnCount) {
        //
        if (cellRange.row == -1 && cellRange.col == -1) {
            return new spreadNS.Range(0, 0, rowCount, columnCount);
        }
        else if (cellRange.row == -1) {
            return new spreadNS.Range(0, cellRange.col, rowCount, cellRange.colCount);
        }
        else if (cellRange.col == -1) {
            return new spreadNS.Range(cellRange.row, 0, cellRange.rowCount, columnCount);
        }

        return cellRange;
    };

    //设置迷你图
    _getElementById("btnAddSparkline").addEventListener('click',function () {
        var sheet = spread.getActiveSheet();

        var range = getActualCellRange(sheet.getSelections()[0], sheet.getRowCount(), sheet.getColumnCount());
        var rc = _getElementById("line_position").value.split(",");
        var r = parseInt(rc[0]);
        var c = parseInt(rc[1]);
        var orientation = parseInt(_getElementById("line_orientation").value);
        var type = parseInt(_getElementById("line_type").value);
        //如果行位置没有输入信息
        if (!isNaN(r) && !isNaN(c)) {
            sheet.setSparkline(r, c, range, orientation, type, setting);
        }
    });

    //删除迷你图
    _getElementById("btnClearSparkline").addEventListener('click',function () {
        var sheet = spread.getActiveSheet();

        var range = getActualCellRange(sheet.getSelections()[0], sheet.getRowCount(), sheet.getColumnCount());

        for (var r = 0; r < range.rowCount; r++) {
            for (var c = 0; c < range.colCount; c++) {
                sheet.removeSparkline(r + range.row, c + range.col);
            }
        }
    });
}

//获取id的方法
function _getElementById(id){
    return document.getElementById(id);
}

//选择选项的方法
function _selectOption(select, value) {
    for (var i = 0; i < select.length; i++) {
        var op = select.options[i];
        if (op.value === value) {
            op.selected = true;
        } else {
            op.selected = false;
        }
    }
}