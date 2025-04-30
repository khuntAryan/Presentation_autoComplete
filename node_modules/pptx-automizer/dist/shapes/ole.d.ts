import { Shape } from '../classes/shape';
import { ImportedElement, ShapeTargetType, Target } from '../types/types';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
export declare class OLEObject extends Shape {
    private readonly oleObjectPath;
    constructor(shape: ImportedElement, targetType: ShapeTargetType, sourceArchive: IArchive);
    private getFileExtension;
    remove(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<OLEObject>;
    prepare(targetTemplate: RootPresTemplate, targetSlideNumber: number, oleObjects?: Target[]): Promise<void>;
    private copyOleObjectFile;
    private appendToContentTypes;
    private updateSlideRels;
    private updateSlideXml;
    private getContentType;
    private removeOleObjectFile;
    private removeFromContentTypes;
    private removeFromSlideRels;
    static getAllOnSlide(archive: IArchive, relsPath: string): Promise<Target[]>;
    modifyOnAddedSlide(targetTemplate: RootPresTemplate, targetSlideNumber: number, oleObjects: Target[]): Promise<void>;
}
