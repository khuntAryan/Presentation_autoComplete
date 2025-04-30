import { ShapeTargetType, SourceIdentifier } from '../types/types';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import HasShapes from './has-shapes';
import { ILayout } from '../interfaces/ilayout';
export declare class Layout extends HasShapes implements ILayout {
    targetType: ShapeTargetType;
    targetMaster: number;
    constructor(params: {
        presentation: IPresentationProps;
        template: PresTemplate;
        sourceIdentifier: SourceIdentifier;
        targetMaster: number;
    });
    /**
     * Appends slideLayout
     * @internal
     * @param targetTemplate
     * @returns append
     */
    append(targetTemplate: RootPresTemplate): Promise<void>;
    /**
     * Copys slide layout files
     * @internal
     */
    copySlideLayoutFiles(): Promise<void>;
    updateRelation(): Promise<void>;
    getName(): Promise<string>;
}
