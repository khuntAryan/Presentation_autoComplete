import { ICounter } from '../interfaces/icounter';
import { RootPresTemplate } from '../interfaces/root-pres-template';
export declare class CountHelper implements ICounter {
    template: RootPresTemplate;
    name: string;
    count: number;
    constructor(name: string, template: RootPresTemplate);
    static increment(name: string, counters: ICounter[]): number | null;
    static count(name: string, counters: ICounter[]): number;
    static reset(counters: ICounter[]): void;
    static getCounterByName(name: string, counters: ICounter[]): ICounter;
    _increment(): number;
    set(): Promise<void>;
    get(): number;
    private calculateCount;
    private static countSlides;
    private static countMasters;
    private static countLayouts;
    private static countThemes;
    private static countCharts;
    private static countOleObjects;
    private static countImages;
}
