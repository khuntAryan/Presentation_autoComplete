import { FileInfo, ArchiveParams, AutomizerFile } from '../types/types';
import IArchive, { ArchivedFolderCallback } from '../interfaces/iarchive';
import { ContentTypeExtension } from '../enums/content-type-map';
export declare class FileHelper {
    static importArchive(file: AutomizerFile, params: ArchiveParams): IArchive;
    static removeFromDirectory(archive: IArchive, dir: string, cb: ArchivedFolderCallback): Promise<string[]>;
    static getFileExtension(filename: string): ContentTypeExtension;
    static getFileInfo(filename: string): FileInfo;
    static check(archive: IArchive, file: string): boolean;
    static isArchive(archive: any): void;
    static fileExistsInArchive(archive: IArchive, file: string): boolean;
    static zipCopyWithRelations(parentClass: any, type: string, sourceNumber: number, targetNumber: number): Promise<void>;
    static zipCopyByIndex(parentClass: any, prefix: any, sourceId: any, targetId: any, suffix?: any): Promise<IArchive>;
    /**
     * Copies a file from one archive to another. The new file can have a different name to the origin.
     * @param {IArchive} sourceArchive - Source archive
     * @param {string} sourceFile - file path and name inside source archive
     * @param {IArchive} targetArchive - Target archive
     * @param {string} targetFile - file path and name inside target archive
     * @return {IArchive} targetArchive as an instance of IArchive
     */
    static zipCopy(sourceArchive: IArchive, sourceFile: string, targetArchive: IArchive, targetFile?: string): Promise<IArchive>;
}
export declare const exists: (dir: string) => boolean;
export declare const makeDirIfNotExists: (dir: string) => void;
export declare const makeDir: (dir: string) => void;
export declare const copyDir: (src: any, dest: any) => Promise<void>;
export declare const ensureDirectoryExistence: (filePath: any) => boolean;
