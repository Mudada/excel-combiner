"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var _this = this;
exports.__esModule = true;
var exceljs_1 = require("exceljs");
var yargs = require("yargs");
var S = require("sanctuary");
var basePath = "/Users/mudada/Code/Script/excel-combiner/excel/";
var outputFp = basePath.concat("output/output.xlsx");
var inputFp = basePath.concat("Tableau Carnot TSN-EP-v3.xlsx");
var workbook = new exceljs_1.Workbook();
var writeWorkbook = function (filePath, workBook) {
    return workbook.xlsx.writeFile(filePath);
};
var readXlsxFile = function (filePath) {
    var workbook = new exceljs_1.Workbook();
    return workbook.xlsx.readFile(filePath);
};
var showSheetsName = function (workbook) {
    workbook.eachSheet(function (wsheet, id) {
        console.log(wsheet.name);
    });
};
var duplicateRow = function (newRow, row) {
    newRow.values = row.values;
};
var rowValues = function (row) {
    var values = [];
    row.eachCell({ includeEmpty: true }, function (c, _) {
        return c.type == 6 ? values.push(c.result) : values.push(c.value);
    });
    return values;
};
var copyStyle = function (baseRow, copiedRow) {
    baseRow.eachCell({ includeEmpty: true }, function (c, cid) {
        c.style = copiedRow.getCell(cid).style;
    });
};
var copyRowUC = function (f, baseWorksheet, copiedWorksheet) {
    baseWorksheet.eachRow({ includeEmpty: true }, function (r, rid) {
        if (f(r)) {
            var aRow = copiedWorksheet.addRow(rowValues(r));
            copyStyle(aRow, r);
        }
    });
};
var copyRow = S.curry3(copyRowUC);
var copyPartRow = function (start, end) { return copyRow(function (r) { return r.number >= end || r.number < start; }); };
var copyFullRow = copyPartRow(0, 0);
var copyPartWorkbookUC = function (start, end, baseWorkbook, wkCopy) {
    baseWorkbook.eachSheet(function (s, sid) {
        console.log("[LOG] Looking at worksheet ".concat(s.name));
        if (wkCopy.getWorksheet(s.name) === undefined) {
            console.log("[LOG] Worksheet ".concat(s.name).concat(" not found in merge file, creating new one"));
            copyPartRow(0, 0)(s)(wkCopy.addWorksheet(s.name));
        }
        else {
            console.log("[LOG] Worksheet".concat(s.name).concat(" found in merge file, adding content"));
            copyPartRow(start, end)(s)(wkCopy.getWorksheet(s.name));
        }
    });
};
var copyPartWorkbook = S.curry4(copyPartWorkbookUC);
var copyWorkbook = function (baseWorkbook, wkCopy) { return copyPartWorkbook(0)(0)(baseWorkbook)(wkCopy); };
/*
  const file = readXlsxFile(inputFp).then(
  (w) => {
  const wkCopy = copyWorkbook(w)
  showSheetsName(wkCopy)
  wkCopy.xlsx.writeFile(outputFp).then(
  (_) => console.log("done")
  ,console.log
  )
  },
  console.log)

*/
var main = function () { return __awaiter(_this, void 0, void 0, function () {
    var argv, wkCopy, _i, _a, file, w, err_1;
    return __generator(this, function (_b) {
        switch (_b.label) {
            case 0:
                argv = yargs
                    .usage('Usage: excombiner -f /dir/where/files/to/combine/are -o output-file')
                    .demandOption(['f', 'o'])
                    .array('f')
                    .argv;
                wkCopy = new exceljs_1.Workbook();
                _i = 0, _a = argv.f;
                _b.label = 1;
            case 1:
                if (!(_i < _a.length)) return [3 /*break*/, 6];
                file = _a[_i];
                console.log("[LOG] Trying to read ".concat(file));
                _b.label = 2;
            case 2:
                _b.trys.push([2, 4, , 5]);
                return [4 /*yield*/, readXlsxFile(file)];
            case 3:
                w = _b.sent();
                console.log("[LOG] Success in reading ".concat(file));
                console.log("[LOG] Mergin ".concat(file));
                copyPartWorkbook(2)(5)(w)(wkCopy);
                console.log("[LOG] ".concat(file).concat(" merged !"));
                return [3 /*break*/, 5];
            case 4:
                err_1 = _b.sent();
                console.log("[ERROR] ".concat(err_1));
                return [3 /*break*/, 5];
            case 5:
                _i++;
                return [3 /*break*/, 1];
            case 6:
                console.log("[LOG] Trying to write merged file to ".concat(argv.o));
                return [4 /*yield*/, wkCopy.xlsx.writeFile(argv.o).then(function (_) {
                        console.log("[LOG] Merge successfull, wrote the file at ".concat(argv.o));
                    }, function (error) { return console.log("[ERROR] ".concat(error)); })];
            case 7:
                _b.sent();
                return [2 /*return*/];
        }
    });
}); };
