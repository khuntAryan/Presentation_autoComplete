import IArchive from '../interfaces/iarchive';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { AnalyzedElementType, AutomizerParams, ElementOnSlide, FindElementSelector, GenerateElements, GenerateOnSlideCallback, ImportedElement, ImportElement, ShapeModificationCallback, ShapeTargetType, SlideModificationCallback, SlidePlaceholder, SourceIdentifier, StatusTracker } from '../types/types';
import { ContentTracker } from '../helper/content-tracker';
import { ElementInfo, XmlDocument, XmlElement } from '../types/xml-types';
import { XmlSlideHelper } from '../helper/xml-slide-helper';
export default class HasShapes {
    /**
     * Source template of slide
     * @internal
     */
    sourceTemplate: PresTemplate;
    /**
     * Target template of slide
     * @internal
     */
    targetTemplate: RootPresTemplate;
    /**
     * Target number of slide
     * @internal
     */
    targetNumber: number;
    /**
     * Source number of slide
     * @internal
     */
    sourceNumber: number;
    /**
     * Target archive of slide
     * @internal
     */
    targetArchive: IArchive;
    /**
     * Source archive of slide
     * @internal
     */
    sourceArchive: IArchive;
    /**
     * Source path of slide
     * @internal
     */
    sourcePath: string;
    /**
     * Target path of slide
     * @internal
     */
    targetPath: string;
    /**
     * Modifications of root template slide
     * @internal
     */
    modifications: SlideModificationCallback[];
    /**
     * Modifications of slide relations
     * @internal
     */
    relModifications: SlideModificationCallback[];
    /**
     * Import elements of slide
     * @internal
     */
    importElements: ImportElement[];
    /**
     * Generate elements on slide with PptxGenJS
     * @internal
     */
    generateElements: GenerateElements[];
    /**
     * Rels path of slide
     * @internal
     */
    relsPath: string;
    /**
     * Target rels path of slide
     * @internal
     */
    targetRelsPath: string;
    /**
     * Root  of slide
     * @internal
     */
    root: IPresentationProps;
    status: StatusTracker;
    content: ContentTracker;
    /**
     * List of unsupported tags in slide xml
     * @internal
     */
    unsupportedTags: string[];
    /**
     * List of unsupported tags in slide xml
     * @internal
     */
    unsupportedRelationTypes: string[];
    targetType: ShapeTargetType;
    params: AutomizerParams;
    cleanupPlaceholders: boolean;
    constructor(params: {
        presentation: IPresentationProps;
        template: PresTemplate;
    });
    /**
     * Asynchronously retrieves all text element IDs from the slide.
     * @returns {Promise<string[]>} A promise that resolves to an array of text element IDs.
     */
    getAllTextElementIds(): Promise<string[]>;
    /**
     * Asynchronously retrieves all elements from the slide.
     * @params filterTags Use an array of strings to filter parent tags (e.g. 'sp')
     * @returns {Promise<ElementInfo[]>} A promise that resolves to an array of ElementInfo objects.
     */
    getAllElements(filterTags?: string[]): Promise<ElementInfo[]>;
    /**
     * Asynchronously retrieves one element from the slide.
     * @params selector Use shape name or creationId to find the shape
     * @returns {Promise<ElementInfo>} A promise that resolves an ElementInfo object.
     */
    getElement(selector: string): Promise<ElementInfo>;
    /**
     * Asynchronously retrieves the dimensions of the slide.
     * This function utilizes the XmlSlideHelper to get the slide dimensions.
     *
     * @returns {Promise<{width: number, height: number}>} A promise that resolves to an object containing the width and height of the slide.
     */
    getDimensions(): Promise<{
        width: number;
        height: number;
    }>;
    /**
     * Asynchronously retrieves an instance of XmlSlideHelper for slide.
     * @returns {Promise<XmlSlideHelper>} An instance of XmlSlideHelper.
     */
    getSlideHelper(): Promise<XmlSlideHelper>;
    /**
     * Push modifications list
     * @internal
     * @param callback
     */
    modify(callback: SlideModificationCallback): void;
    /**
     * Push relations modifications list
     * @internal
     * @param callback
     */
    modifyRelations(callback: SlideModificationCallback): void;
    /**
     * Select and modify a single element on an added slide.
     * @param {string} selector - Element's name on the slide.
     * Should be a unique string defined on the "Selection"-pane within ppt.
     * @param {ShapeModificationCallback | ShapeModificationCallback[]} callback - One or more callback functions to apply.
     * Depending on the shape type (e.g. chart or table), different arguments will be passed to the callback.
     */
    modifyElement(selector: FindElementSelector, callback: ShapeModificationCallback | ShapeModificationCallback[]): this;
    generate(generate: GenerateOnSlideCallback, objectName?: string): this;
    getGeneratedElements(): GenerateElements[];
    /**
     * Select, insert and (optionally) modify a single element to a slide.
     * @param {string} presName - Filename or alias name of the template presentation.
     * Must have been importet with Automizer.load().
     * @param {number} slideNumber - Slide number within the specified template to search for the required element.
     * @param {FindElementSelector} selector - a string or object to find the target element
     * @param {ShapeModificationCallback | ShapeModificationCallback[]} callback - One or more callback functions to apply.
     * Depending on the shape type (e.g. chart or table), different arguments will be passed to the callback.
     */
    addElement(presName: string, slideNumber: number, selector: FindElementSelector, callback?: ShapeModificationCallback | ShapeModificationCallback[]): this;
    /**
     * Remove a single element from slide.
     * @param {string} selector - Element's name on the slide.
     */
    removeElement(selector: FindElementSelector): this;
    /**
     * Adds element to modifications list
     * @internal
     * @param presName
     * @param slideNumber
     * @param selector
     * @param mode
     * @param [callback]
     * @returns element to modifications list
     */
    private addElementToModificationsList;
    /**
     * ToDo: Implement creationIds as well for slideMasters
     *
     * Try to convert a given slide's creationId to corresponding slide number.
     * Used if automizer is run with useCreationIds: true
     * @internal
     * @param PresTemplate
     * @slideNumber SourceSlideIdentifier
     * @returns number
     */
    getSlideNumber(template: PresTemplate, slideIdentifier: SourceIdentifier): number;
    /**
     * Imported selected elements
     * @internal
     */
    importedSelectedElements(): Promise<void>;
    /**
     * Gets element info
     * @internal
     * @param importElement
     * @returns element info
     */
    getElementInfo(importElement: ImportElement): Promise<ImportedElement>;
    /**
     * @param selector
     * @param sourceArchive
     * @param sourcePath
     * @param useCreationIds
     */
    findElementOnSlide(selector: FindElementSelector, sourceArchive: IArchive, sourcePath: string, useCreationIds: boolean): Promise<ElementOnSlide>;
    checkIntegrity(info: boolean, assert: boolean): Promise<void>;
    /**
     * Adds slide to presentation
     * @internal
     * @returns slide to presentation
     */
    addToPresentation(): Promise<void>;
    /**
     * Appends to slide rel
     * @internal
     * @param rootArchive
     * @param relId
     * @param slideCount
     * @returns to slide rel
     */
    appendToSlideRel(rootArchive: IArchive, relId: string, slideCount: number): Promise<XmlElement>;
    /**
     * Appends a new slide to slide list in presentation.xml.
     * If rootArchive has no slides, a new node will be created.
     * "id"-attribute of 'p:sldId'-element must be greater than 255.
     * @internal
     * @param rootArchive
     * @param relId
     * @returns to slide list
     */
    appendToSlideList(rootArchive: IArchive, relId: string): Promise<XmlElement>;
    /**
     * Appends a new slide to slide list in presentation.xml.
     * If rootArchive has no slides, a new node will be created.
     * "id"-attribute of 'p:sldId'-element must be greater than 255.
     * @internal
     * @param rootArchive
     * @param relId
     * @returns to slide list
     */
    appendToSlideMasterList(rootArchive: IArchive, relId: string): Promise<XmlElement>;
    /**
     * Appends slide to content type
     * @internal
     * @param rootArchive
     * @param slideCount
     * @returns slide to content type
     */
    appendToContentType(rootArchive: IArchive, count: number): Promise<XmlElement>;
    /**
     * slideNote numbers differ from slide numbers if presentation
     * contains slides without notes. We need to find out
     * the proper enumeration of slideNote xml files.
     * @internal
     * @returns slide note file number
     */
    getSlideNoteSourceNumber(): Promise<number | undefined>;
    /**
     * Copys slide note files
     * @internal
     * @returns slide note files
     */
    copySlideNoteFiles(sourceNotesNumber: number): Promise<void>;
    /**
     * Updates slide note file
     * @internal
     * @returns slide note file
     */
    updateSlideNoteFile(sourceNotesNumber: number): Promise<void>;
    /**
     * Appends notes to content type
     * @internal
     * @param rootArchive
     * @param slideCount
     * @returns notes to content type
     */
    appendNotesToContentType(rootArchive: IArchive, slideCount: number): Promise<XmlElement>;
    /**
     * Copys related content
     * @internal
     * @returns related content
     */
    copyRelatedContent(): Promise<void>;
    /**
     * Analyzes element
     * @internal
     * @param sourceElement
     * @param sourceArchive
     * @param slideNumber
     * @returns element
     */
    analyzeElement(sourceElement: XmlElement, sourceArchive: IArchive, relsPath: string): Promise<AnalyzedElementType>;
    private findHyperlinkInElement;
    /**
     * Applys modifications
     * @internal
     * @returns modifications
     */
    applyModifications(): Promise<void>;
    /**
     * Apply modifications to slide relations
     * @internal
     * @returns modifications
     */
    applyRelModifications(): Promise<void>;
    /**
     * Removes all unsupported tags from slide xml.
     * E.g. added relations & tags by Thinkcell cannot
     * be processed by pptx-automizer at the moment.
     * @internal
     */
    cleanSlide(targetPath: string, sourcePlaceholderTypes?: SlidePlaceholder[]): Promise<void>;
    /**
     * If you insert a placeholder shape on a target slide with an empty
     * placeholder of the same type, we need to remove the existing
     * placeholder.
     *
     * @param xml
     * @param sourcePlaceholderTypes
     */
    removeDuplicatePlaceholders(xml: XmlDocument, sourcePlaceholderTypes: SlidePlaceholder[]): void;
    /**
     * If a placeholder shape was inserted on a slide without a corresponding
     * placeholder, powerPoint will usually smash the shape's formatting.
     * This function removes the placeholder tag.
     * @param xml
     * @param sourcePlaceholderTypes
     */
    normalizePlaceholderShapes(xml: XmlDocument, sourcePlaceholderTypes: SlidePlaceholder[]): void;
    /**
     * Removes all unsupported relations from _rels xml.
     * @internal
     */
    cleanRelations(targetRelsPath: string): Promise<void>;
    parsePlaceholders(): Promise<SlidePlaceholder[]>;
}
