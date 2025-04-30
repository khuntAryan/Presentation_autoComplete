import { ImportedElement, ShapeModificationCallback, ShapeTargetType } from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { Shape } from '../classes/shape';
import { XmlElement } from '../types/xml-types';
export declare class GenericShape extends Shape {
    sourceElement: XmlElement;
    callbacks: ShapeModificationCallback[];
    constructor(shape: ImportedElement, targetType: ShapeTargetType);
    modify(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<GenericShape>;
    append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<GenericShape>;
    remove(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<GenericShape>;
    prepare(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<void>;
}
