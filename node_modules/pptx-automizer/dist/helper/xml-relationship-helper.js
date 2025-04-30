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
Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlRelationshipHelper = void 0;
const xml_helper_1 = require("./xml-helper");
const general_helper_1 = require("./general-helper");
const file_helper_1 = require("./file-helper");
const crypto_1 = require("crypto");
class XmlRelationshipHelper {
    constructor(xml, tag) {
        this.xmlTargets = [];
        this.targets = [];
        if (xml) {
            this.setXml(xml);
        }
        this.tag = tag || 'Relationship';
        return this;
    }
    initialize(archive, file, path, prefix) {
        return __awaiter(this, void 0, void 0, function* () {
            this.archive = archive;
            this.file = file;
            this.path = path + '/';
            const fileProxy = yield this.archive;
            this.xml = yield xml_helper_1.XmlHelper.getXmlFromArchive(fileProxy, this.path + this.file);
            yield this.readTargets();
            if (prefix) {
                return this.getTargetsByPrefix(prefix);
            }
            return this;
        });
    }
    setXml(xml) {
        this.xml = xml;
        return this;
    }
    getTargetsByPrefix(prefix) {
        const prefixes = typeof prefix === 'string' ? [prefix] : prefix;
        const targets = [];
        this.xmlTargets.forEach((xmlTarget) => {
            prefixes.forEach((prefix) => {
                const target = XmlRelationshipHelper.parseRelationTarget(xmlTarget, prefix, true);
                if (target === null || target === void 0 ? void 0 : target.prefix) {
                    targets.push(target);
                }
            });
        });
        return targets;
    }
    getTargetsByType(type) {
        const targets = [];
        this.xmlTargets.forEach((xmlTarget) => {
            const target = XmlRelationshipHelper.parseRelationTarget(xmlTarget);
            if ((target === null || target === void 0 ? void 0 : target.type) === type) {
                targets.push(target);
            }
        });
        return targets;
    }
    getTargetByRelId(findRid) {
        const matchedTarget = this.xmlTargets.find((xmlTarget) => xmlTarget.getAttribute('Id') === findRid);
        if (matchedTarget) {
            return XmlRelationshipHelper.parseRelationTarget(matchedTarget);
        }
    }
    readTargets() {
        if (this.xmlTargets.length) {
            return this;
        }
        const relationshipItems = this.xml.getElementsByTagName(this.tag);
        for (const i in relationshipItems) {
            if (relationshipItems[i] &&
                relationshipItems[i].getAttribute !== undefined) {
                this.xmlTargets.push(relationshipItems[i]);
            }
        }
        return this;
    }
    /**
     * This will copy all unhandled related contents into
     * the target archive.
     *
     * Pptx messages on opening a corrupted file are most likely
     * caused by broken relations and this is going to prevent
     * files from being missed.
     *
     * @param sourceArchive
     * @param check
     * @param assert
     */
    assertRelatedContent(sourceArchive, check, assert) {
        return __awaiter(this, void 0, void 0, function* () {
            for (const xmlTarget of this.xmlTargets) {
                const targetFile = xmlTarget.getAttribute('Target');
                const targetMode = xmlTarget.getAttribute('TargetMode');
                const targetPath = targetFile.replace('../', 'ppt/');
                if (targetMode !== 'External' &&
                    this.archive.fileExists(targetPath) === false) {
                    // ToDo: There are falsy errors on files that have already been
                    //       copied with another target name.
                    // if (check) {
                    //   if (typeof sourceArchive.filename === 'string') {
                    //     console.error(
                    //       'Related content from ' +
                    //         sourceArchive.filename +
                    //         ' not found: ' +
                    //         targetFile,
                    //     );
                    //   } else {
                    //     console.error('Related content not found: ' + targetFile);
                    //   }
                    // }
                    if (assert) {
                        const target = XmlRelationshipHelper.parseRelationTarget(xmlTarget);
                        const buf = (0, crypto_1.randomBytes)(5).toString('hex');
                        const targetSuffix = '-' + buf + '.' + target.filenameExt;
                        yield file_helper_1.FileHelper.zipCopy(sourceArchive, targetPath, this.archive, targetPath + targetSuffix);
                        xmlTarget.setAttribute('Target', targetFile + targetSuffix);
                        yield xml_helper_1.XmlHelper.appendImageExtensionToContentType(this.archive, target.filenameExt);
                    }
                }
            }
        });
    }
    static parseRelationTarget(element, prefix, matchByPrefix) {
        if (!element || element.getAttribute === undefined) {
            return;
        }
        const type = element.getAttribute('Type');
        const file = element.getAttribute('Target');
        const rId = element.getAttribute('Id');
        const filename = (0, general_helper_1.last)(file.split('/'));
        const relType = (0, general_helper_1.last)(type.split('/'));
        const filenameExt = (0, general_helper_1.last)(filename.split('.'));
        const filenameMatch = filename
            .replace('.' + filenameExt, '')
            .match(/^(.+?)(\d+)*$/);
        const number = filenameMatch && filenameMatch[2] ? Number(filenameMatch[2]) : 0;
        const filenameBase = filenameMatch && filenameMatch[1] ? filenameMatch[1] : filename;
        const target = {
            rId,
            type,
            file,
            filename,
            relType,
            element,
            filenameExt,
            filenameMatch,
            number,
            filenameBase,
            getTargetValue: () => target.element.getAttribute('Target'),
            updateTargetValue: (newTarget) => {
                target.element.setAttribute('Target', newTarget);
            },
            updateId: (newId) => {
                target.element.setAttribute('Id', newId);
            },
        };
        if (prefix) {
            const subtype = (0, general_helper_1.last)(prefix.split('/'));
            if (matchByPrefix &&
                !XmlRelationshipHelper.targetMatchesRelationship(relType, subtype, file, prefix)) {
                return;
            }
            return this.extendTarget(prefix, subtype, target);
        }
        return target;
    }
    static extendTarget(prefix, subtype, target) {
        return Object.assign(Object.assign({}, target), { prefix,
            subtype, updateTargetIndex: (newIndex) => {
                target.element.setAttribute('Target', `${prefix}${newIndex}.xml`);
            } });
    }
    static targetMatchesRelationship(relType, subtype, target, prefix) {
        if (relType === 'package')
            return true;
        return relType === subtype && target.indexOf(prefix) === 0;
    }
    static getSlideLayoutNumber(sourceArchive, slideId) {
        return __awaiter(this, void 0, void 0, function* () {
            const slideToLayouts = yield new XmlRelationshipHelper().initialize(sourceArchive, `slide${slideId}.xml.rels`, `ppt/slides/_rels`, '../slideLayouts/slideLayout');
            return slideToLayouts[0].number;
        });
    }
    static getSlideMasterNumber(sourceArchive, slideLayoutId) {
        return __awaiter(this, void 0, void 0, function* () {
            const layoutToMaster = (yield new XmlRelationshipHelper().initialize(sourceArchive, `slideLayout${slideLayoutId}.xml.rels`, `ppt/slideLayouts/_rels`, '../slideMasters/slideMaster'));
            return layoutToMaster[0].number;
        });
    }
}
exports.XmlRelationshipHelper = XmlRelationshipHelper;
//# sourceMappingURL=xml-relationship-helper.js.map