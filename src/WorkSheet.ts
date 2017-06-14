import * as XLSX from "xlsx";
import * as utils from "./utils/utils"

export class WorkSheet{

    private sheetJson = null;
    private me:XLSX.WorkSheet = null;

    public constructor(sheet:XLSX.WorkSheet){
        this.me = sheet;
        this.sheetJson = XLSX.utils.sheet_to_json(this.me,{header:1});
    }

    public getRowNum():number{
        
        var rowNum = this.sheetJson.length;
        return rowNum;
    }

    public getColumnNum(rowNum:number):number{
        var row = this.sheetJson[rowNum];
        return row.length;
    }

    public getCellValue(rowNum:number, columnNum:any):any{
        var cell = new utils.Cell(rowNum, utils.translateExcelColNumToString(columnNum));
        return this.me[cell.toString()].v;
    }

    public getCellAddress(rowNum:number,columnNum:number):utils.Cell{
        var l:string = utils.translateExcelColNumToString(columnNum);
        return new utils.Cell(rowNum, l);
    }

    public addCell(cell:utils.Cell, content:any){
        var t:string, v:any, r:string, h:string, w:string;
        v = content;
        h = content.toString();
        w = content.toString();
        if(typeof content == "number"){
            t = 'n' ; r = '<t>'+content.toString()+'</t>'; 
            this.me[cell.toString()] ={t:t,v:v, r:r, h:h, w:w};
        }else{
            t = 's'; 
            this.me[cell.toString()] = {t:t, v:v, w:w};
        }
        //update !ref
        this.updateRef(cell);
        
    }

    private updateRef(cell:utils.Cell):void{
        var ref:string = this.me["!ref"];
        var refEnd = ref.substr(ref.indexOf(":")+1);
        var refEndRow = refEnd.substr(0, refEnd.match("[0-9]").index);
        if(cell.r>new Number(refEndRow).valueOf()){
            refEndRow = cell.r.toString();
        }
        var refEndColumn = refEnd.substr(refEnd.match("[0-9]").index+1);
        if(cell.c > refEndColumn){
            refEndColumn = cell.c.toString();
        }
        this.me["!ref"] = ref.substr(0, ref.indexOf(":")+1)+refEndColumn+refEndRow;
    }
}