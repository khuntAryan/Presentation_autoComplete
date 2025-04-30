/// <reference types="node" />
/// <reference types="node" />
import Archive from './archive';
import JSZip from 'jszip';
import { AutomizerFile, AutomizerParams } from '../../types/types';
import IArchive, { ArchivedFile } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
export default class ArchiveJszip extends Archive implements IArchive {
    archive: JSZip;
    file: Buffer;
    constructor(filename: AutomizerFile);
    private initialize;
    fileExists(file: string): boolean;
    folder(dir: string): Promise<ArchivedFile[]>;
    read(file: string, type: 'string' | 'nodebuffer'): Promise<string | Buffer>;
    write(file: string, data: string | Buffer): Promise<this>;
    remove(file: string): Promise<void>;
    extract(file: string): Promise<ArchiveJszip>;
    output(location: string, params: AutomizerParams): Promise<void>;
    stream(params: AutomizerParams, options?: JSZip.JSZipGeneratorOptions<'nodebuffer'>): Promise<NodeJS.ReadableStream>;
    getFinalArchive(): Promise<JSZip>;
    getContent(params: AutomizerParams): Promise<Buffer>;
    readXml(file: string): Promise<XmlDocument>;
    writeXml(file: string, XmlDocument: XmlDocument): void;
}
