"use strict";
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
exports.OLEObject = void 0;
const file_helper_1 = require("../helper/file-helper");
const xml_helper_1 = require("../helper/xml-helper");
const shape_1 = require("../classes/shape");
const path_1 = __importDefault(require("path"));
class OLEObject extends shape_1.Shape {
    constructor(shape, targetType, sourceArchive) {
        var _a;
        super(shape, targetType);
        this.sourceArchive = sourceArchive;
        this.oleObjectPath = `ppt/embeddings/${this.sourceRid}${this.getFileExtension((_a = shape.target) === null || _a === void 0 ? void 0 : _a.file)}`;
        this.relRootTag = 'p:oleObj';
        this.relAttribute = 'r:id';
    }
    getFileExtension(file) {
        if (!file)
            return '.bin';
        const ext = path_1.default.extname(file).toLowerCase();
        return ['.bin', '.xls', '.xlsx', '.doc', '.docx', '.ppt', '.pptx'].includes(ext)
            ? ext
            : '.bin';
    }
    // NOTE: modify() and append() won't be implemented.
    // TODO: remove is not currently properly implemented,
    //  suggest we delete the file from the archive as well as removing the relationship.
    remove(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.removeFromSlideTree();
            yield this.removeOleObjectFile();
            yield this.removeFromContentTypes();
            yield this.removeFromSlideRels();
            return this;
        });
    }
    prepare(targetTemplate, targetSlideNumber, oleObjects) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.setTarget(targetTemplate, targetSlideNumber);
            const allOleObjects = oleObjects ||
                (yield OLEObject.getAllOnSlide(this.sourceArchive, this.targetSlideRelFile));
            const oleObject = allOleObjects.find((obj) => obj.rId === this.sourceRid);
            if (!oleObject) {
                throw new Error(`OLE object with rId ${this.sourceRid} not found.`);
            }
            const sourceFilePath = `ppt/embeddings/${oleObject.file.split('/').pop()}`;
            this.createdRid = yield xml_helper_1.XmlHelper.getNextRelId(this.targetArchive, this.targetSlideRelFile);
            yield this.copyOleObjectFile(sourceFilePath);
            yield this.appendToContentTypes();
            yield this.updateSlideRels();
            yield this.updateSlideXml();
        });
    }
    copyOleObjectFile(sourceFilePath) {
        return __awaiter(this, void 0, void 0, function* () {
            const fileExtension = this.getFileExtension(sourceFilePath);
            const targetFileName = `ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
            try {
                yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, sourceFilePath, this.targetArchive, targetFileName);
            }
            catch (error) {
                console.error('Error copying OLE object file:', error);
                throw error;
            }
        });
    }
    appendToContentTypes() {
        return __awaiter(this, void 0, void 0, function* () {
            const contentTypesPath = '[Content_Types].xml';
            const contentTypesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, contentTypesPath);
            const types = contentTypesXml.getElementsByTagName('Types')[0];
            const fileExtension = this.getFileExtension(this.oleObjectPath);
            const partName = `/ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
            const existingOverride = Array.from(types.getElementsByTagName('Override')).find((override) => override.getAttribute('PartName') === partName);
            if (!existingOverride) {
                const newOverride = contentTypesXml.createElement('Override');
                newOverride.setAttribute('PartName', partName);
                newOverride.setAttribute('ContentType', this.getContentType(fileExtension));
                types.appendChild(newOverride);
                yield xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, contentTypesPath, contentTypesXml);
            }
        });
    }
    updateSlideRels() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
            const relationships = relXml.getElementsByTagName('Relationship');
            const fileExtension = this.getFileExtension(this.oleObjectPath);
            const newTarget = `../embeddings/oleObject${this.createdRid}${fileExtension}`;
            // Update or create the relationship
            let relationshipUpdated = false;
            for (let i = 0; i < relationships.length; i++) {
                if (relationships[i].getAttribute('Id') === this.sourceRid) {
                    relationships[i].setAttribute('Id', this.createdRid);
                    relationships[i].setAttribute('Target', newTarget);
                    relationshipUpdated = true;
                    break;
                }
            }
            if (!relationshipUpdated) {
                const newRel = relXml.createElement('Relationship');
                newRel.setAttribute('Id', this.createdRid);
                newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject');
                newRel.setAttribute('Target', newTarget);
                relXml.documentElement.appendChild(newRel);
            }
            yield xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
        });
    }
    updateSlideXml() {
        return __awaiter(this, void 0, void 0, function* () {
            const slideXmlPath = `ppt/slides/slide${this.targetSlideNumber}.xml`;
            const slideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, slideXmlPath);
            const oleObjs = Array.from(slideXml.getElementsByTagName('p:oleObj'));
            oleObjs.forEach((oleObj) => {
                if (oleObj.getAttribute('r:id') === this.sourceRid) {
                    oleObj.setAttribute('r:id', this.createdRid);
                    const oleObjPr = oleObj.getElementsByTagName('p:oleObjPr')[0];
                    if (oleObjPr) {
                        const links = Array.from(oleObjPr.getElementsByTagName('a:link'));
                        links.forEach((link) => {
                            link.setAttribute('r:id', this.createdRid);
                        });
                    }
                }
            });
            yield xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, slideXmlPath, slideXml);
        });
    }
    getContentType(fileExtension) {
        const contentTypes = {
            '.bin': 'application/vnd.openxmlformats-officedocument.oleObject',
            '.xls': 'application/vnd.ms-excel',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.doc': 'application/msword',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.ppt': 'application/vnd.ms-powerpoint',
            '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        };
        return (contentTypes[fileExtension.toLowerCase()] ||
            'application/vnd.openxmlformats-officedocument.oleObject');
    }
    removeOleObjectFile() {
        return __awaiter(this, void 0, void 0, function* () {
            const fileExtension = this.getFileExtension(this.oleObjectPath);
            const fileName = `ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
            yield this.targetArchive.remove(fileName);
        });
    }
    removeFromContentTypes() {
        return __awaiter(this, void 0, void 0, function* () {
            const contentTypesPath = '[Content_Types].xml';
            const contentTypesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, contentTypesPath);
            const types = contentTypesXml.getElementsByTagName('Types')[0];
            const fileExtension = this.getFileExtension(this.oleObjectPath);
            const partName = `/ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
            const overrideToRemove = Array.from(types.getElementsByTagName('Override')).find((override) => override.getAttribute('PartName') === partName);
            if (overrideToRemove) {
                types.removeChild(overrideToRemove);
                yield xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, contentTypesPath, contentTypesXml);
            }
        });
    }
    removeFromSlideRels() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
            const relationships = relXml.getElementsByTagName('Relationship');
            for (let i = 0; i < relationships.length; i++) {
                if (relationships[i].getAttribute('Id') === this.createdRid) {
                    relationships[i].parentNode.removeChild(relationships[i]);
                    break;
                }
            }
            yield xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
        });
    }
    static getAllOnSlide(archive, relsPath) {
        return __awaiter(this, void 0, void 0, function* () {
            const oleObjectType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject';
            return xml_helper_1.XmlHelper.getRelationshipItems(archive, relsPath, (element, rels) => {
                const type = element.getAttribute('Type');
                if (type === oleObjectType) {
                    rels.push({
                        rId: element.getAttribute('Id'),
                        type: element.getAttribute('Type'),
                        file: element.getAttribute('Target'),
                        element: element,
                    });
                }
            });
        });
    }
    modifyOnAddedSlide(targetTemplate, targetSlideNumber, oleObjects) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber, oleObjects);
        });
    }
}
exports.OLEObject = OLEObject;
//# sourceMappingURL=ole.js.map