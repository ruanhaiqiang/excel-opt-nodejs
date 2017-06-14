export function translateExcelColNumToString(n){
	var intPart = Math.floor(n/26);
	var remainderPart = n%26;
	if(remainderPart ==0){
		remainderPart=26;
		intPart--;
	}
	var lastString = String.fromCharCode(65+remainderPart-1);
	if(intPart>0){
		var intString = translateExcelColNumToString(intPart);
		return intString+lastString;
	}else{
		return lastString;
	}
}

export class Cell{
    r:number;
    c:string;
    public constructor(r:number, c:string){
        this.r = r;
        this.c = c;
    }
    public toString(){
        return this.c+this.r;
    }
}