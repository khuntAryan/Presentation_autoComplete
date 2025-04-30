"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.compressFolder = void 0;
const fs_1 = __importStar(require("fs"));
const path_1 = __importDefault(require("path"));
const jszip_1 = __importDefault(require("jszip"));
// Thanks to https://github.com/DesignByOnyx
// see https://github.com/Stuk/jszip/issues/386 for more info
/**
 * Returns a flat list of all files and subfolders for a directory (recursively).
 * @param {string} dir
 * @returns {Promise<string[]>}
 */
const getFilePathsRecursively = (dir) => __awaiter(void 0, void 0, void 0, function* () {
    // returns a flat array of absolute paths of all files recursively contained in the dir
    const list = yield fs_1.promises.readdir(dir);
    const statPromises = list.map((file) => __awaiter(void 0, void 0, void 0, function* () {
        const fullPath = path_1.default.resolve(dir, file);
        const stat = yield fs_1.promises.stat(fullPath);
        if (stat && stat.isDirectory()) {
            return getFilePathsRecursively(fullPath);
        }
        return fullPath;
    }));
    return (yield Promise.all(statPromises)).flat(Infinity);
});
/**
 * Creates an in-memory zip stream from a folder in the file system
 * @param {string} dir
 * @returns {JSZip}
 */
const createZipFromFolder = (dir) => __awaiter(void 0, void 0, void 0, function* () {
    const absRoot = path_1.default.resolve(dir);
    const filePaths = yield getFilePathsRecursively(dir);
    return filePaths.reduce((z, filePath) => {
        const relative = filePath.replace(absRoot, '');
        // create folder trees manually :(
        const zipFolder = path_1.default
            .dirname(relative)
            .split(path_1.default.sep)
            .reduce((zf, dirName) => zf.folder(dirName), z);
        zipFolder.file(path_1.default.basename(filePath), fs_1.default.createReadStream(filePath));
        return z;
    }, new jszip_1.default());
});
/**
 * Compresses a folder to the specified zip file.
 * @param {string} srcDir
 * @param {string} destFile
 */
const compressFolder = (srcDir, destFile, options) => __awaiter(void 0, void 0, void 0, function* () {
    const start = Date.now();
    try {
        const zip = yield createZipFromFolder(srcDir);
        zip
            .generateNodeStream(Object.assign({ streamFiles: true }, options))
            .pipe(fs_1.default.createWriteStream(destFile))
            .on('error', (err) => console.error('Error writing file', err.stack))
            .on('finish', () => console.log('Zip written successfully:', Date.now() - start, 'ms'));
    }
    catch (ex) {
        console.error('Error creating zip', ex);
    }
});
exports.compressFolder = compressFolder;
//# sourceMappingURL=jszip-helper.js.map