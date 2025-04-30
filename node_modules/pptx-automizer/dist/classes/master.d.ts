import { ShapeTargetType, SourceIdentifier, Target } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { IMaster } from '../interfaces/imaster';
import HasShapes from './has-shapes';
export declare class Master extends HasShapes implements IMaster {
    targetType: ShapeTargetType;
    key: string;
    constructor(params: {
        presentation: IPresentationProps;
        template: PresTemplate;
        sourceIdentifier: SourceIdentifier;
    });
    static getKey(slideLayoutNumber: number, templateName: string): string;
    /**
     * Appends slide
     * @internal
     * @param targetTemplate
     * @returns append
     */
    append(targetTemplate: RootPresTemplate): Promise<void>;
    copyRelatedLayouts(): Promise<Target[]>;
    copyThemeFiles(): Promise<void>;
    /**
     * Copy slide master files
     * @internal
     */
    copySlideMasterFiles(): Promise<void>;
    appendThemeToContentType(rootArchive: IArchive, themeCount: string | number): Promise<XmlElement>;
}
