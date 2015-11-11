'use strict';
var gutil = require('gulp-util');
var through = require('through2');
var XLSX = require('xlsx');


var toJson = function (fileName, headRow, valueRow) {
    var workbook;
    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }
    var sheet_name_list = workbook.SheetNames;

    var worksheet = workbook.Sheets[sheet_name_list[0]];

    var namemap = [];

    var json = {};

    var curRow = 0;
    for (var key in worksheet) {

        if (worksheet.hasOwnProperty(key)) {
            var cell = worksheet[key];
            var match = /([A-Z]+)(\d)/.exec(key);
            if (!match) {
                continue;
            }
            var col = match[1]; // ABCD
            var row = match[2]; // 1234
            var value = cell.v;


            if (row == headRow) {
                namemap[col] = value;
            } else if (row < valueRow) {
                //continue;
            } else {
                if (col == "A") {
                    json[cell.v] = {};
                    curRow = cell.v;
                }
                json[curRow][namemap[col]] = cell.v;
            }
        }
    }
    return json;
};


module.exports = function (options) {
    options = options || {};
    return through.obj(function (file, enc, cb) {
        // 如果文件为空，不做任何操作，转入下一个操作，即下一个 .pipe()
        if (file.isNull()) {
            this.push(file);
            return cb();
        }

        // 插件不支持对 Stream 对直接操作，跑出异常
        if (file.isStream()) {
            this.emit('error', new gutil.PluginError(PLUGIN_NAME, 'Streaming not supported'));
            return cb();
        }

        // 然后将处理后的字符串，再转成Buffer形式
        //var content = pp.preprocess(file.contents.toString(), options || {});
        //file.contents = new Buffer(content);
        var arr = [];
        for (var i = 0; i < file.contents.length; ++i) arr[i] = String.fromCharCode(file.contents[i]);
        var bString = arr.join("");

        /* Call XLSX */
        var workbook = XLSX.read(bString, {type: "binary"});
        file.contents = new Buffer(JSON.stringify(toJson(workbook, options.headRow || 1, options.valueRowStart || 2)));

        if (options.trace) {
            console.log("convert file :" + file.path);
        }

        // 下面这两句基本是标配啦，可以参考下 through2 的API
        this.push(file);
        cb();
    });
};
