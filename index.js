'use strict';
var gutil = require('gulp-util');
var through = require('through2');
var XLSX = require('xlsx');

/**
 * excel filename or workbook to json
 * @param fileName
 * @param headRow
 * @param valueRow
 * @returns {{}} json
 */
var toJson = function (fileName, headRow, valueRow) {
    var workbook;
    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }
    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var namemap = [];

    // json to return
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
        if (file.isNull()) {
            this.push(file);
            return cb();
        }

        if (file.isStream()) {
            this.emit('error', new gutil.PluginError(PLUGIN_NAME, 'Streaming not supported'));
            return cb();
        }

        var arr = [];
        for (var i = 0; i < file.contents.length; ++i) arr[i] = String.fromCharCode(file.contents[i]);
        var bString = arr.join("");

        /* Call XLSX */
        var workbook = XLSX.read(bString, {type: "binary"});
        file.contents = new Buffer(JSON.stringify(toJson(workbook, options.headRow || 1, options.valueRowStart || 2)));

        if (options.trace) {
            console.log("convert file :" + file.path);
        }

        this.push(file);
        cb();
    });
};
