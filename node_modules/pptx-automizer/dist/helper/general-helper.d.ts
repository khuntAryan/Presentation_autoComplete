export declare class GeneralHelper {
    static arrayify<T>(input: T | T[]): T[];
    static propertyExists<T>(object: T, property: string): boolean;
}
export declare const vd: (v: any, keys?: boolean) => void;
export declare const last: <T>(arr: T[]) => T;
export interface Logger {
    verbosity: 0 | 1 | 2;
    target: 'console' | 'file';
    log: (message: string, verbosity: Logger['verbosity'], showStack?: boolean, target?: Logger['target']) => void;
}
export declare const Logger: Logger;
export declare const log: (message: string, verbosity: Logger['verbosity']) => void;
export declare const logDebug: (message: string, verbosity: Logger['verbosity']) => void;
