import { Target } from '../types/types';
import { SlideInfo, TemplateSlideInfo, XmlDocument, XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
export declare class XmlTemplateHelper {
    archive: IArchive;
    relType: string;
    relTypeNotes: string;
    path: string;
    defaultSlideName: string;
    constructor(archive: IArchive);
    getCreationIds(): Promise<SlideInfo[]>;
    parseSlideRelFile(slideRelFile: string): number;
    getSlideInfo(slideXml: XmlDocument, archive: any, slideRelFile: string): Promise<TemplateSlideInfo>;
    getNameFromSlideInfo(slideXml: XmlDocument): string;
    getSlideNoteRels(archive: IArchive, slideRelFile: string): Promise<Target[]>;
    getSlideNameFromNotes(archive: any, slideNoteRels: any): Promise<string>;
    parseTitleElement(titleElement: XmlElement): string[];
    /**
     * Returns the slide numbers of a given template as a sorted array of integers.
     * @returns {Promise<number[]>} - A promise that resolves to a sorted array of slide numbers in the template.
     */
    getAllSlideNumbers(): Promise<number[]>;
}
