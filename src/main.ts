import { Workbook, Worksheet } from 'exceljs'
import * as S from 'sanctuary'

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
    = basePath.concat("/output/output.xlsx")

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

const main = async () => {
    const file = readXlsxFile(inputFp).then(
        showSheetsName,
        console.log)
}

main()
