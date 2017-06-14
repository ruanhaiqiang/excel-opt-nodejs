///

import * as XLSX from "xlsx";
import * as utils from "./utils/utils"
import * as WorkSheet from "./WorkSheet";

export class WorkBook{
    private sheets :null;
    private sheetNames :null;
    public workbook :XLSX.WorkBook;


    public constructor(workbook:XLSX.WorkBook){
        this.workbook = workbook;
    }

    public getSheetByName(sheetName:string):WorkSheet.WorkSheet{
        var sheet = this.workbook.Sheets[sheetName];
        var worksheet = new WorkSheet.WorkSheet(sheet);
        return worksheet;
    }

    public getSheetNames():string[]{
        return this.workbook.SheetNames;
    }
}