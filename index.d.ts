declare module 'excel-opt-nodejs'{

    export interface IExcelOpt{
        /**
         * @param filename
         * @return
         *      if filename is not exist and format error, return null
         *      else return xlsx.workbook
         */
        readFile(filename:string, opts?: IParsingOptions);

        writeFile(workbook:any, filename, opts?: IWritingOptions);
    } 

    export interface IWorkBook{
        getSheetByName(sheetName:string):IWorkSheet;

        getSheetNames():string[];
    }

    export interface IWorkSheet{
        /**
         * get sheet row number
         */
        getRowNum():number;
        /**
         * get column number
         * @param rowNum :number
         */
        getColumnNum(rowNum:number):number;
        /**
         * get cell value
         */
        getCellValue(rowNum:number, columnNum:any):any;

        /**
         * get cell address
         */
        getCellAddress(rowNum:number,columnNum:number):Cell;

        /** 
         * add cell to sheet
         */
        addCell(cell:Cell, content:any);

    }
    export interface Cell{
        toString();
    }
    export interface IParsingOptions {
        /**
         * Input data encoding
         */
        type?: 'base64' | 'binary' | 'buffer' | 'array' | 'file';

        /**
         * Save formulae to the .f field
         * @default true
         */
        cellFormula?: boolean;

        /**
         * Parse rich text and save HTML to the .h field
         * @default true
         */
        cellHTML?: boolean;

        /**
         * Save number format string to the .z field
         * @default false
         */
        cellNF?: boolean;

        /**
         * Save style/theme info to the .s field
         * @default false
         */
        cellStyles?: boolean;

        /**
         * Store dates as type d (default is n)
         * @default false
         */
        cellDates?: boolean;

        /**
         * Create cell objects for stub cells
         * @default false
         */
        sheetStubs?: boolean;

        /**
         * If >0, read the first sheetRows rows
         * @default 0
         */
        sheetRows?: number;

        /**
         * If true, parse calculation chains
         * @default false
         */
        bookDeps?: boolean;

        /**
         * If true, add raw files to book object
         * @default false
         */
        bookFiles?: boolean;

        /**
         * If true, only parse enough to get book metadata
         * @default false
         */
        bookProps?: boolean;

        /**
         * If true, only parse enough to get the sheet names
         * @default false
         */
        bookSheets?: boolean;

        /**
         * If true, expose vbaProject.bin to vbaraw field
         * @default false
         */
        bookVBA?: boolean;

        /**
         * If defined and file is encrypted, use password
         * @default ''
         */
        password?: string;
    }

    export interface IWritingOptions {
        /**
         * Output data encoding
         */
        type?: 'base64' | 'binary' | 'buffer' | 'file';

        /**
         * Store dates as type d (default is n)
         * @default false
         */
        cellDates?: boolean;

        /**
         * Generate Shared String Table
         * @default false
         */
        bookSST?: boolean;

        /**
         * Type of Workbook
         * @default 'xlsx'
         */
        bookType?: 'xlsx' | 'xlsm' | 'xlsb' | 'ods' | 'biff2' | 'fods' | 'csv';

        /**
         * Name of Worksheet for single-sheet formats
         * @default ''
         */
        sheet?: string;

        /**
         * Use ZIP compression for ZIP-based formats
         * @default false
         */
        compression?: boolean;
    }

}