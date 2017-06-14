import * as XLSX from "xlsx";
import * as utils from "./utils/utils"
import * as WorkBook from "./WorkBook";

export class ExcelOpt{

    public readFile(filename:string,opts?:object):WorkBook.WorkBook{
        var workbook = XLSX.readFile(filename, opts);
        return new WorkBook.WorkBook(workbook);
    }

    public writeFile(workBook:WorkBook.WorkBook, filename:string, opts?:object){
        XLSX.writeFile(workBook.workbook, filename, opts);
    }
}





