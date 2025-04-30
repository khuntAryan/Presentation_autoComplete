import { XmlDocument, XmlElement } from '../types/xml-types';
import { Target } from '../types/types';
import IArchive from '../interfaces/iarchive';
import { ElementSubtype } from '../enums/element-type';
export declare class XmlRelationshipHelper {
    archive: IArchive;
    file: string;
    path: string;
    tag: string;
    xml: XmlDocument;
    xmlTargets: XmlElement[];
    targets: Target[];
    constructor(xml?: XmlDocument, tag?: string);
    initialize(archive: IArchive, file: string, path: string, prefix?: string): Promise<Target[] | this>;
    setXml(xml: any): this;
    getTargetsByPrefix(prefix: string | string[]): Target[];
    getTargetsByType(type: string): Target[];
    getTargetByRelId(findRid: string): Target | null;
    readTargets(): this;
    /**
     * This will copy all unhandled related contents into
     * the target archive.
     *
     * Pptx messages on opening a corrupted file are most likely
     * caused by broken relations and this is going to prevent
     * files from being missed.
     *
     * @param sourceArchive
     * @param check
     * @param assert
     */
    assertRelatedContent(sourceArchive: IArchive, check?: boolean, assert?: boolean): Promise<void>;
    static parseRelationTarget(element: XmlElement, prefix?: string, matchByPrefix?: boolean): Target | undefined;
    static extendTarget(prefix: string, subtype: ElementSubtype, target: Target): Target;
    static targetMatchesRelationship(relType: any, subtype: any, target: any, prefix: any): boolean;
    static getSlideLayoutNumber(sourceArchive: any, slideId: number): Promise<any>;
    static getSlideMasterNumber(sourceArchive: any, slideLayoutId: number): Promise<number>;
}
