type PrettyPrintToken = {
    match: string;
    tag: string;
    offset: number;
    preContent: string;
};
export declare class XmlPrettyPrint {
    xmlStr: string;
    TAB: string;
    constructor(xmlStr: string);
    dump(): void;
    prettify(): string;
    parse(xmlStr: string): string[];
    getToken(regex: RegExp, str: string): PrettyPrintToken;
    addLine(output: string[], content: string, indent: number): void;
}
export {};
