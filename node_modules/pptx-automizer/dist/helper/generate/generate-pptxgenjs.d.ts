import PptxGenJS from 'pptxgenjs';
import { ISlide } from '../../interfaces/islide';
import Automizer from '../../automizer';
import { GenerateElements } from '../../types/types';
import { IGenerator } from '../../interfaces/igenerator';
import { IPptxGenJSSlide } from '../../interfaces/ipptxgenjs-slide';
/**
 * Using pptxGenJs on an automizer ISlide will create a temporary pptx template
 * and auto-import the generated shapes to the right place on the output slides.
 */
export default class GeneratePptxGenJs implements IGenerator {
    tmpFile: string;
    slides: ISlide[];
    generator: PptxGenJS;
    automizer: Automizer;
    countSlides: number;
    constructor(automizer: Automizer, slides: ISlide[]);
    create(): void;
    generateSlides(): Promise<void>;
    generateElements(generate: GenerateElements[], pgenSlide: any, tmpSlideNumber: any): Promise<void>;
    addElements(generate: GenerateElements[], slide: ISlide): void;
    /**
     * This is a wrapper around supported pptxGenJS slide item types.
     * It is required to create a unique objectName and find the generated
     * shapes by object name later.
     *
     * @param pgenSlide
     * @param generateElement
     * @param addedObjects
     */
    addSlideItems: (pgenSlide: PptxGenJS.Slide, generateElement: GenerateElements, addedObjects: string[]) => IPptxGenJSSlide;
    generateObjectName(generateElement: GenerateElements, addedObjects: string[]): string;
    getOptions: (options: any, objectName: string) => any;
    appendPptxGenSlide(): PptxGenJS.Slide;
    cleanup(): Promise<void>;
}
