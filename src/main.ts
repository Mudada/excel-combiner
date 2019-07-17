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

const writeWorkbook = async (filePath: string, workBook: Workbook): Promise<Either<string, void>> => {
    try {
        await workbook.xlsx.writeFile(filePath)
        return S.Right (undefined)
    } catch (error) {
        return S.Left (error)
    }
}

const readXlsxFile = async (filePath: string): Promise<Either<string, Workbook>> => {
    try {
        const workbook = new Workbook()
        await workbook.xlsx.readFile(filePath)
        return S.Right (workbook)
    } catch (error) {
        return S.Left (error)
    }
}

const showSheetsName = (workbook: Workbook):void => {
    workbook.eachSheet((wsheet: Worksheet, id: number) => {
        console.log(wsheet.name)
    })
}

const main = async () => {
    const file = await readXlsxFile(inputFp)
    S.either ((err: string):void => {console.log(err)}) (showSheetsName) (file)
}

main()
