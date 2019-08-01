import { Cell, Row, Workbook, Worksheet } from 'exceljs'
import * as yargs from 'yargs'
import * as S from 'sanctuary'
import * as fs from 'fs'
import * as jsonfile from 'jsonfile'

interface Maybe<A> {
    constructor: {
        '@@type': 'sanctuary/Maybe';
    };
}

interface Either<A, B> {
    constructor: {
        '@@type': 'sanctuary/Either';
    };
}

const basePath: string
    = "/Users/mudada/Code/Script/excel-combiner/excel/"

const outputFp: string
    = basePath.concat("output/output.xlsx")

const inputFp: string
    = basePath.concat("Tableau Carnot TSN-EP-v3.xlsx")

const workbook: Workbook
    = new Workbook()

const writeWorkbook = (filePath: string, workBook: Workbook): Promise<void> => {
        return workbook.xlsx.writeFile(filePath)
}

const readXlsxFile = (filePath: string): Promise<Workbook> => {
    const workbook = new Workbook()
    return workbook.xlsx.readFile(filePath)
}

const showSheetsName = (workbook: Workbook):void => {
    workbook.eachSheet((wsheet: Worksheet, id: number) => {
        console.log(wsheet.name)
    })
}

const duplicateRow = (newRow: Row, row: Row):void => {
    newRow.values = row.values
}

const rowValues = (row: Row):string[] => {
    const values = []
    row.eachCell({includeEmpty: true}, (c,_) =>
                 c.type == 6 ? values.push(c.result) : values.push(c.value))
    return values
}

const copyStyle = (baseRow: Row, copiedRow: Row):void => {
    baseRow.eachCell({includeEmpty: true}, (c, cid) => {
        c.style = copiedRow.getCell(cid).style
    })
}

const copyRowUC = (f:(row: Row) => boolean,
                       baseWorksheet: Worksheet,
                       copiedWorksheet: Worksheet):void => {
    baseWorksheet.eachRow({includeEmpty: true}, (r, rid) => {
        if(f(r)) {
            const aRow = copiedWorksheet.addRow(rowValues(r))
            copyStyle(aRow, r)
        }
    })
}

const copyRow = S.curry3 (copyRowUC)

const copyPartRow = (start: number, end: number) => copyRow ((r: Row) => r.number >= end || r.number < start)

const copyFullRow = copyPartRow(0, 0)

const copyPartWorkbookUC = (start: number,
                            end: number,
                            baseWorkbook: Workbook,
                            wkCopy: Workbook):void => {
    baseWorkbook.eachSheet((s, sid) => {
        console.log("[LOG] Looking at worksheet ".concat(s.name))
        if (wkCopy.getWorksheet(s.name) === undefined) {
            console.log("[LOG] Worksheet ".concat(s.name).concat(" not found in merge file, creating new one"))
            copyFullRow (s) (wkCopy.addWorksheet(s.name))
        }
        else {
            console.log("[LOG] Worksheet ".concat(s.name).concat(" found in merge file, adding content"))
            copyPartRow (start, end) (s) (wkCopy.getWorksheet(s.name))
        }
    })
}

const copyPartWorkbook = S.curry4(copyPartWorkbookUC)

const copyWorkbook = (baseWorkbook: Workbook
                      , wkCopy: Workbook) => copyPartWorkbook (0) (0) (baseWorkbook) (wkCopy)


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


interface swOpt { name: string, except: string[] }
type SWOPT = swOpt | string

const isString = (val: any): val is string => {
    return typeof val === 'string'
}

const formatWorkbookForSWOPTUC = (wk: Workbook, swopt: SWOPT):void => {
    if (isString(swopt))
        wk.eachSheet((s, _) => s.name = swopt)
    else
        wk.eachSheet((s, _) => S.elem (s.name) (swopt.except) ? s.name : s.name = swopt.name)

}

const formatWorkbookForSWOPT = S.curry2(formatWorkbookForSWOPTUC)

const parseSingleWorksheetOpt = (opt: any): Either<string, SWOPT> => {
    if (isString(opt))
        return S.Right (opt)
    if (opt.name !== undefined
        && isString(opt.name)
        && opt.except !== undefined
        && Array.isArray(opt.except)
       )
        return S.Right ({name: opt.name, except: opt.except})
    return S.Left ("[ERROR] JSON config file, singleWorksheet option wrongly defined.")
}

const main = async () => {
    const argv = yargs
        .usage('Usage: excombiner -f /dir/where/files/to/combine/are -o output-file')
        .demandOption(['f'])
        .argv

    console.log(argv.f)
    jsonfile.readFile(argv.f).then(
        async (conf) => {
            const wkCopy = new Workbook()

            for (let file of conf.files) {
                console.log("[LOG] Trying to read ".concat(file))
                try {
                    const w = await readXlsxFile(file)
                    console.log("[LOG] Success in reading ".concat(file))
                    console.log("[LOG] Mergin ".concat(file))
                    if (conf.singleWorksheet !== undefined) {
                        console.log("[LOG] Single worksheet option selected, using ".concat(conf.singleWorksheet))
                        const swOpt = parseSingleWorksheetOpt(conf.singleWorksheet)
                        S.either (console.log) (formatWorkbookForSWOPT(w)) (swOpt)
                    }
                    copyPartWorkbook (conf.start) (conf.end) (w) (wkCopy)
                    console.log("[LOG] ".concat(file).concat(" merged !"))
                } catch (err) {
                    console.log("[ERROR] ".concat(err))
                }
            }

            console.log("[LOG] Trying to write merged file to ".concat(argv.o))

            await wkCopy.xlsx.writeFile(conf.output).then(
                (_) => {
                    console.log("[LOG] Merge successfull, wrote the file at ".concat(argv.o))
                }
                , (error) => console.log("[ERROR] ".concat(error))
            )
        },
        console.log
    )
}
main()
