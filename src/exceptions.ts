export class CellModifyingError extends Error {
    constructor(public message: string){
        super(message);
        this.name = "Cell Error"
    }
}

export class ClearNonCellElementValueError extends Error {
    constructor(public message: string){
        super(message);
        this.name = "Non Cell Element Modification Error"
    }
}