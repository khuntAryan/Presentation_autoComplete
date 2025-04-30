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
exports.Master = void 0;
const file_helper_1 = require("../helper/file-helper");
const xml_helper_1 = require("../helper/xml-helper");
const xml_relationship_helper_1 = require("../helper/xml-relationship-helper");
const has_shapes_1 = __importDefault(require("./has-shapes"));
const layout_1 = require("./layout");
const general_helper_1 = require("../helper/general-helper");
class Master extends has_shapes_1.default {
    constructor(params) {
        super(params);
        this.targetType = 'slideMaster';
        // ToDo analogue for slideMasters
        // this.sourceNumber = this.getSlideNumber(
        //   params.template,
        //   params.sourceIdentifier,
        // );
        this.sourceNumber = Number(params.sourceIdentifier);
        this.key = Master.getKey(this.sourceNumber, params.template.name);
        this.sourcePath = `ppt/slideMasters/slideMaster${this.sourceNumber}.xml`;
        this.relsPath = `ppt/slideMasters/_rels/slideMaster${this.sourceNumber}.xml.rels`;
    }
    static getKey(slideLayoutNumber, templateName) {
        return slideLayoutNumber + '@' + templateName;
    }
    /**
     * Appends slide
     * @internal
     * @param targetTemplate
     * @returns append
     */
    append(targetTemplate) {
        return __awaiter(this, void 0, void 0, function* () {
            this.targetTemplate = targetTemplate;
            this.targetArchive = yield targetTemplate.archive;
            this.targetNumber = targetTemplate.incrementCounter('masters');
            this.targetPath = `ppt/slideMasters/slideMaster${this.targetNumber}.xml`;
            this.targetRelsPath = `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`;
            this.sourceArchive = yield this.sourceTemplate.archive;
            (0, general_helper_1.log)('Importing slideMaster ' + this.targetNumber, 2);
            yield this.copySlideMasterFiles();
            yield this.copyRelatedLayouts();
            yield this.copyRelatedContent();
            yield this.addToPresentation();
            yield this.copyThemeFiles();
            if (this.importElements.length) {
                yield this.importedSelectedElements();
            }
            yield this.applyModifications();
            yield this.applyRelModifications();
            const info = this.targetTemplate.automizer.params.showIntegrityInfo;
            const assert = this.targetTemplate.automizer.params.showIntegrityInfo;
            yield this.checkIntegrity(info, assert);
            yield this.cleanSlide(this.targetPath);
        });
    }
    copyRelatedLayouts() {
        return __awaiter(this, void 0, void 0, function* () {
            const targets = (yield new xml_relationship_helper_1.XmlRelationshipHelper().initialize(this.targetArchive, `slideMaster${this.targetNumber}.xml.rels`, `ppt/slideMasters/_rels`, '../slideLayouts/slideLayout'));
            for (const target of targets) {
                const layout = new layout_1.Layout({
                    presentation: this.targetTemplate.automizer,
                    template: this.sourceTemplate,
                    sourceIdentifier: target.number,
                    targetMaster: this.targetNumber,
                });
                yield this.targetTemplate.appendLayout(layout);
                const layoutName = yield layout.getName();
                this.targetTemplate.mapContents('slideLayout', this.sourceTemplate.name, target.number, layout.targetNumber, layoutName);
                target.updateTargetIndex(layout.targetNumber);
            }
            return targets;
        });
    }
    copyThemeFiles() {
        return __awaiter(this, void 0, void 0, function* () {
            const targets = yield xml_helper_1.XmlHelper.getRelationshipTargetsByPrefix(this.targetArchive, `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`, '../theme/theme');
            if (!targets.length) {
                return;
            }
            const themeTarget = targets[0];
            const themeSourceId = themeTarget.number;
            const themeTargetId = this.targetTemplate.incrementCounter('themes');
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/theme/theme${themeSourceId}.xml`, this.targetArchive, `ppt/theme/theme${themeTargetId}.xml`);
            yield this.appendThemeToContentType(this.targetArchive, themeTargetId);
            yield xml_helper_1.XmlHelper.replaceAttribute(this.targetArchive, `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`, 'Relationship', 'Id', themeTarget.rId, `../theme/theme${themeTargetId}.xml`, 'Target');
        });
    }
    /**
     * Copy slide master files
     * @internal
     */
    copySlideMasterFiles() {
        return __awaiter(this, void 0, void 0, function* () {
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/slideMasters/slideMaster${this.sourceNumber}.xml`, this.targetArchive, `ppt/slideMasters/slideMaster${this.targetNumber}.xml`);
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/slideMasters/_rels/slideMaster${this.sourceNumber}.xml.rels`, this.targetArchive, `ppt/slideMasters/_rels/slideMaster${this.targetNumber}.xml.rels`);
        });
    }
    appendThemeToContentType(rootArchive, themeCount) {
        return xml_helper_1.XmlHelper.append(xml_helper_1.XmlHelper.createContentTypeChild(rootArchive, {
            PartName: `/ppt/theme/theme${themeCount}.xml`,
            ContentType: `application/vnd.openxmlformats-officedocument.theme+xml`,
        }));
    }
}
exports.Master = Master;
//# sourceMappingURL=master.js.map