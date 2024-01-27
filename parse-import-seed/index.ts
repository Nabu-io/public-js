import validateExcelFile from "./validateFile"
import { processWorkbook } from "./processWorkbook"

export default (file: File): Promise<any> => {

  return validateExcelFile(file)
    .then((workbook: any) => processWorkbook(workbook))
    .catch((error: any) => {
      // do something
    })

}