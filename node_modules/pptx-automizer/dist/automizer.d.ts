/// <reference types="node" />
import { ArchiveParams, AutomizerFile, AutomizerParams, AutomizerSummary, PresentationInfo, SourceIdentifier, StatusTracker } from './types/types';
import { IPresentationProps } from './interfaces/ipresentation-props';
import { PresTemplate } from './interfaces/pres-template';
import { RootPresTemplate } from './interfaces/root-pres-template';
import { ModifyXmlCallback, TemplateInfo } from './types/xml-types';
import { ContentTracker } from './helper/content-tracker';
import JSZip from 'jszip';
import { ISlide } from './interfaces/islide';
import { IMaster } from './interfaces/imaster';
/**
 * Automizer
 *
 * The basic class for `pptx-automizer` package.
 * This class will be exported as `Automizer` by `index.ts`.
 */
export default class Automizer implements IPresentationProps {
    rootTemplate: RootPresTemplate;
    /**
     * Templates  of automizer
     * @internal
     */
    templates: PresTemplate[];
    templateDir: string;
    templateFallbackDir: string;
    outputDir: string;
    archiveParams: ArchiveParams;
    /**
     * Timer of automizer
     * @internal
     */
    timer: number;
    params: AutomizerParams;
    status: StatusTracker;
    content: ContentTracker;
    modifyPresentation: ModifyXmlCallback[];
    /**
     * Creates an instance of `pptx-automizer`.
     * @param [params]
     */
    constructor(params: AutomizerParams);
    setStatusTracker(statusTracker: StatusTracker['next']): void;
    /**
  
     */
    presentation(): Promise<this>;
    /**
     * Load a pptx file and set it as root template.
     * @param file - Filename, path to the template or Buffer containing the file.
     * Filenames and paths will be prefixed with 'templateDir'
     * @returns Instance of Automizer
     */
    loadRoot(file: AutomizerFile): this;
    /**
     * Load a template pptx file.
     * @param file - Filename, path to the template or Buffer containing the file.
     * Filenames and paths will be prefixed with 'templateDir'
     * @param name - Optional short name for a template loaded from a file. If skipped, the template will be named by its location.
     * if the file is a Buffer the name is required.
     * @returns Instance of Automizer
     */
    load(file: AutomizerFile, name?: string): this;
    /**
     * Loads a pptx file either as a root template as a template file.
     * A name can be specified to give templates an alias.
     * @param location
     * @param [name]
     * @returns template
     */
    private loadTemplate;
    /**
     * Load media files to output presentation.
     * @returns Instance of Automizer
     * @param filename Filename or path to the media file.
     * @param dir Specify custom path for media instead of mediaDir from AutomizerParams.
     */
    loadMedia(filename: string | string[], dir?: string, prefix?: string): this;
    /**
     * Parses all loaded templates and collects creationIds for slides and
     * elements. This will make finding templates and elements independent
     * of slide number and element name.
     * @returns Promise<TemplateInfo[]>
     */
    setCreationIds(): Promise<TemplateInfo[]>;
    /**
     * Get some info about the imported templates
     * @returns Promise<PresentationInfo>
     */
    getInfo(): Promise<PresentationInfo>;
    /**
     * Determines whether template is root or default template.
     * @param template
     * @returns pres template
     */
    private isPresTemplate;
    /**
     * Add a slide from one of the imported templates by slide number or creationId.
     * @param name - Name or alias of the template; must have been loaded with `Automizer.load()`
     * @param slideIdentifier - Number or creationId of slide in template presentation
     * @param callback - Executed after slide was added. The newly created slide will be passed to the callback as first argument.
     * @returns Instance of Automizer
     */
    addSlide(name: string, slideIdentifier: SourceIdentifier, callback?: (slide: ISlide) => void): this;
    /**
     * Copy and modify a master and the associated layouts from template to output.
     *
     * @param name
     * @param sourceIdentifier
     * @param callback
     */
    addMaster(name: string, sourceIdentifier: number, callback?: (slideMaster: IMaster) => void): this;
    /**
     * Searches this.templates to find template by given name.
     * @internal
     * @param name Alias name if given to loaded template.
     * @returns template
     */
    getTemplate(name: string): PresTemplate;
    /**
     * Write all imports and modifications to a file.
     * @param location - Filename or path for the file. Will be prefixed with 'outputDir'
     * @returns summary object.
     */
    write(location: string): Promise<AutomizerSummary>;
    /**
     * Create a ReadableStream from output pptx file.
     * @param generatorOptions - JSZipGeneratorOptions for nodebuffer Output type
     * @returns Promise<NodeJS.ReadableStream>
     */
    stream(generatorOptions?: JSZip.JSZipGeneratorOptions<'nodebuffer'>): Promise<NodeJS.ReadableStream>;
    /**
     * Pass final JSZip instance.
     * @returns Promise<NodeJS.ReadableStream>
     */
    getJSZip(): Promise<JSZip>;
    finalizePresentation(): Promise<void>;
    /**
     * Write all masterSlides to archive.
     */
    writeMasterSlides(): Promise<void>;
    /**
     * Write all slides to archive.
     */
    writeSlides(): Promise<void>;
    /**
     * Write all media files to archive.
     */
    writeMediaFiles(): Promise<void>;
    /**
     * Applies all callbacks in this.modifyPresentation-array.
     * The callback array can be pushed by this.modify()
     */
    applyModifyPresentationCallbacks(): Promise<void>;
    /**
     * Apply some callbacks to restore archive/xml structure
     * and prevent corrupted pptx files.
     *
     * TODO: Use every imported image only once
     * TODO: Check for lost relations
     */
    normalizePresentation(): Promise<void>;
    modify(cb: ModifyXmlCallback): this;
    /**
     * Applies path prefix to given location string.
     * @param location path and/or filename
     * @param [type] template or output
     * @returns location
     */
    private getLocation;
}
