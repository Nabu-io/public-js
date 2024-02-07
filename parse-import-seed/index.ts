import validateExcelFile from "./validateFile"
import processWorkbook from "./processWorkbook"

export default function parseImportSeed(file: ArrayBuffer): Promise<any> {
  return validateExcelFile(file)
    .then((workbook: any) => processWorkbook(workbook))
    .catch((error: any) => {
      // do something
    })
}