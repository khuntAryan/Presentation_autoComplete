import { ICounter } from '../interfaces/icounter';
import { ISlide } from '../interfaces/islide';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ITemplate } from '../interfaces/itemplate';
import { ContentMap, SlideInfo } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { ArchiveParams, AutomizerFile, MediaFile } from '../types/types';
import Automizer from '../automizer';
import { IMaster } from '../interfaces/imaster';
import { ILayout } from '../interfaces/ilayout';
import { IGenerator } from '../interfaces/igenerator';
export declare class Template implements ITemplate {
    /**
     * Path to local file
     * @type string
     */
    location: string;
    /**
     * An alias name to identify template and simplify
     * @type string
     */
    name: string;
    /**
     * Node file buffer
     * @type InputType
     */
    file: any;
    /**
     * this.file will be passed to FileProxy
     * @type Archive
     */
    archive: IArchive;
    /**
     * Array containing all slides coming from Automizer.addSlide()
     * @type: ISlide[]
     */
    slides: ISlide[];
    /**
     * Array containing all slideMasters coming from Automizer.addMaster()
     * @type: IMaster[]
     */
    masters: IMaster[];
    /**
     * Array containing all counters
     * @type: ICounter[]
     */
    counter: ICounter[];
    creationIds: SlideInfo[];
    slideNumbers: number[];
    existingSlides: number;
    contentMap: ContentMap[];
    mediaFiles: MediaFile[];
    automizer: Automizer;
    generator: IGenerator;
    constructor(file: AutomizerFile, params: ArchiveParams);
    static import(file: AutomizerFile, params: ArchiveParams, automizer?: Automizer): PresTemplate | RootPresTemplate;
    mapContents(type: 'slideMaster' | 'slideLayout', key: string, sourceId: number, targetId: number, name?: string): void;
    getNamedMappedContent(type: 'slideMaster' | 'slideLayout', name: string): ContentMap;
    getMappedContent(type: 'slideMaster' | 'slideLayout', key: string, sourceId: number): ContentMap;
    /**
     * Returns the slide numbers of a given template as a sorted array of integers.
     * @returns {Promise<number[]>} - A promise that resolves to a sorted array of slide numbers in the template.
     */
    getAllSlideNumbers(): Promise<number[]>;
    setCreationIds(): Promise<SlideInfo[]>;
    appendMasterSlide(slideMaster: IMaster): Promise<void>;
    appendSlide(slide: ISlide): Promise<void>;
    appendLayout(slideLayout: ILayout): Promise<void>;
    countExistingSlides(): Promise<void>;
    truncate(): Promise<void>;
    getSlideIdList(): Promise<Document>;
    initializeCounter(): Promise<void>;
    incrementCounter(name: string): number;
    count(name: string): number;
    runExternalGenerator(): Promise<void>;
    cleanupExternalGenerator(): Promise<void>;
}
