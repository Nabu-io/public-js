import validateExcelFile from "./validateFile"
import processWorkbook from "./processWorkbook"

export default function parseImportSeed(file: ArrayBuffer): Promise<any> {
  return validateExcelFile(file)
    .then((workbook: any) => processWorkbook(workbook))
    .catch((error: any) => {
      console.error('Error parsing import seed: ', error)
      throw error
    })
}