export default class CellIdHelper {
    private _chars;
    private _nextId;
    constructor(chars?: string);
    start(index: number): this;
    next(): string;
    _increment(): void;
    [Symbol.iterator](): Generator<string, void, unknown>;
    static increment(letterNumber: number): string;
    static setRange(range: string, colId: number, length?: number): string;
    static getSpanString(startColNumber: number, startRowNumber: number, cols: number, rows: number): string;
    static getCellAddressString(c: number, r: number): string;
}
