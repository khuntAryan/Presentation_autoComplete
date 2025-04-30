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
exports.contentTracker = exports.ContentTracker = void 0;
const file_helper_1 = require("./file-helper");
const xml_helper_1 = require("./xml-helper");
const constants_1 = require("../constants/constants");
class ContentTracker {
    constructor() {
        this.files = {
            'ppt/slideMasters': [],
            'ppt/slideLayouts': [],
            'ppt/slides': [],
            'ppt/charts': [],
            'ppt/embeddings': [],
        };
        this.relations = {
            // '.': [],
            'ppt/slides/_rels': [],
            'ppt/slideMasters/_rels': [],
            'ppt/slideLayouts/_rels': [],
            'ppt/charts/_rels': [],
            'ppt/_rels': [],
            ppt: [],
        };
        this.relationTags = (0, constants_1.contentTrack)();
    }
    reset() {
        ['files', 'relations'].forEach((section) => Object.keys(this[section]).forEach((key) => {
            this[section][key] = [];
        }));
        this.relationTags = (0, constants_1.contentTrack)();
    }
    trackFile(file) {
        const info = file_helper_1.FileHelper.getFileInfo(file);
        if (this.files[info.dir]) {
            this.files[info.dir].push(info.base);
        }
    }
    trackRelation(file, attributes) {
        const info = file_helper_1.FileHelper.getFileInfo(file);
        if (this.relations[info.dir]) {
            this.relations[info.dir].push({
                base: info.base,
                attributes,
            });
        }
    }
    analyzeContents(archive) {
        return __awaiter(this, void 0, void 0, function* () {
            this.setArchive(archive);
            yield this.analyzeRelationships();
            yield this.trackSlideMasters();
            yield this.trackSlideLayouts();
        });
    }
    setArchive(archive) {
        this.archive = archive;
    }
    /**
     * This will be replaced by future slideMaster handling.
     */
    trackSlideMasters() {
        return __awaiter(this, void 0, void 0, function* () {
            const slideMasters = this.getRelationTag('ppt/presentation.xml').getTrackedRelations('slideMaster');
            yield this.addAndAnalyze(slideMasters, 'ppt/slideMasters');
        });
    }
    trackSlideLayouts() {
        return __awaiter(this, void 0, void 0, function* () {
            const usedSlideLayouts = this.getRelationTag('ppt/slideMasters').getTrackedRelations('slideLayout');
            yield this.addAndAnalyze(usedSlideLayouts, 'ppt/slideLayouts');
        });
    }
    addAndAnalyze(trackedRelations, section) {
        return __awaiter(this, void 0, void 0, function* () {
            const targets = yield this.getRelatedContents(trackedRelations);
            targets.forEach((target) => {
                this.trackFile(section + '/' + target.filename);
            });
            const relationTagInfo = this.getRelationTag(section);
            yield this.analyzeRelationship(relationTagInfo);
        });
    }
    getRelatedContents(trackedRelations) {
        return __awaiter(this, void 0, void 0, function* () {
            const relatedContents = [];
            for (const trackedRelation of trackedRelations) {
                for (const target of trackedRelation.targets) {
                    const trackedRelationInfo = yield target.getRelatedContent();
                    relatedContents.push(trackedRelationInfo);
                }
            }
            return relatedContents;
        });
    }
    getRelationTag(source) {
        return exports.contentTracker.relationTags.find((relationTag) => relationTag.source === source);
    }
    analyzeRelationships() {
        return __awaiter(this, void 0, void 0, function* () {
            for (const relationTagInfo of this.relationTags) {
                yield this.analyzeRelationship(relationTagInfo);
            }
        });
    }
    analyzeRelationship(relationTagInfo) {
        return __awaiter(this, void 0, void 0, function* () {
            relationTagInfo.getTrackedRelations = (role) => {
                return relationTagInfo.tags.filter((tag) => tag.role === role);
            };
            for (const relationTag of relationTagInfo.tags) {
                relationTag.targets = relationTag.targets || [];
                if (relationTagInfo.isDir === true) {
                    const files = this.files[relationTagInfo.source] || [];
                    if (!files.length) {
                        // vd('no files');
                        // vd(relationTagInfo.source);
                    }
                    for (const file of files) {
                        yield this.pushRelationTagTargets(relationTagInfo.source + '/' + file, file, relationTag, relationTagInfo);
                    }
                }
                else {
                    const pathInfo = file_helper_1.FileHelper.getFileInfo(relationTagInfo.source);
                    yield this.pushRelationTagTargets(relationTagInfo.source, pathInfo.base, relationTag, relationTagInfo);
                }
            }
        });
    }
    pushRelationTagTargets(file, filename, relationTag, relationTagInfo) {
        return __awaiter(this, void 0, void 0, function* () {
            const attribute = relationTag.attribute || 'r:id';
            const addTargets = yield xml_helper_1.XmlHelper.getRelationshipItems(this.archive, file, (element, rels) => {
                rels.push({
                    file,
                    filename,
                    rId: element.getAttribute(attribute),
                    type: relationTag.type,
                });
            }, relationTag.tag);
            this.addCreatedRelationsFunctions(addTargets, exports.contentTracker.relations[relationTagInfo.relationsKey], relationTagInfo);
            relationTag.targets = [...relationTag.targets, ...addTargets];
        });
    }
    addCreatedRelationsFunctions(addTargets, createdRelations, relationTagInfo) {
        addTargets.forEach((addTarget) => {
            addTarget.getCreatedContent = this.getCreatedContent(createdRelations, addTarget);
            addTarget.getRelatedContent = this.addRelatedContent(relationTagInfo, addTarget);
        });
    }
    getCreatedContent(createdRelations, addTarget) {
        return () => {
            return createdRelations.find((relation) => {
                var _a;
                return (relation.base === addTarget.filename + '.rels' &&
                    ((_a = relation.attributes) === null || _a === void 0 ? void 0 : _a.Id) === addTarget.rId);
            });
        };
    }
    addRelatedContent(relationTagInfo, addTarget) {
        return () => __awaiter(this, void 0, void 0, function* () {
            if (addTarget.relatedContent)
                return addTarget.relatedContent;
            const relationsFile = relationTagInfo.isDir === true
                ? relationTagInfo.relationsKey + '/' + addTarget.filename + '.rels'
                : relationTagInfo.relationsKey;
            const relationTarget = yield xml_helper_1.XmlHelper.getRelationshipItems(this.archive, relationsFile, (element, rels) => {
                const rId = element.getAttribute('Id');
                if (rId === addTarget.rId) {
                    const target = element.getAttribute('Target');
                    const targetMode = element.getAttribute('TargetMode');
                    const fileInfo = file_helper_1.FileHelper.getFileInfo(target);
                    if (targetMode !== 'External') {
                        rels.push({
                            file: target,
                            filename: fileInfo.base,
                            rId: rId,
                            type: element.getAttribute('Type'),
                        });
                    }
                }
            });
            addTarget.relatedContent = relationTarget.find((relationTarget) => relationTarget.rId === addTarget.rId);
            return addTarget.relatedContent;
        });
    }
    collect(section, role, collection) {
        return __awaiter(this, void 0, void 0, function* () {
            collection = collection || [];
            const trackedRelationTag = this.getRelationTag(section);
            const trackedRelations = trackedRelationTag.getTrackedRelations(role);
            const relatedTargets = yield this.getRelatedContents(trackedRelations);
            relatedTargets.forEach((relatedTarget) => collection.push(relatedTarget.filename));
            return collection;
        });
    }
    filterRelations(section, target) {
        const relations = this.relations[section];
        return relations.filter((rel) => rel.attributes.Target === target);
    }
}
exports.ContentTracker = ContentTracker;
exports.contentTracker = new ContentTracker();
//# sourceMappingURL=content-tracker.js.map