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
exports.Layout = void 0;
const file_helper_1 = require("../helper/file-helper");
const xml_helper_1 = require("../helper/xml-helper");
const xml_relationship_helper_1 = require("../helper/xml-relationship-helper");
const has_shapes_1 = __importDefault(require("./has-shapes"));
const general_helper_1 = require("../helper/general-helper");
class Layout extends has_shapes_1.default {
    constructor(params) {
        super(params);
        this.targetType = 'slideLayout';
        this.sourceNumber = Number(params.sourceIdentifier);
        this.targetMaster = params.targetMaster;
        this.sourcePath = `ppt/slideLayouts/slideLayout${this.sourceNumber}.xml`;
        this.relsPath = `ppt/slideLayouts/_rels/slideLayout${this.sourceNumber}.xml.rels`;
    }
    /**
     * Appends slideLayout
     * @internal
     * @param targetTemplate
     * @returns append
     */
    append(targetTemplate) {
        return __awaiter(this, void 0, void 0, function* () {
            this.targetTemplate = targetTemplate;
            this.targetArchive = yield targetTemplate.archive;
            this.targetNumber = targetTemplate.incrementCounter('layouts');
            this.targetPath = `ppt/slideLayouts/slideLayout${this.targetNumber}.xml`;
            this.targetRelsPath = `ppt/slideLayouts/_rels/slideLayout${this.targetNumber}.xml.rels`;
            this.sourceArchive = yield this.sourceTemplate.archive;
            (0, general_helper_1.log)('Importing slideLayout ' + this.targetNumber, 2);
            yield this.copySlideLayoutFiles();
            yield this.copyRelatedContent();
            yield this.addToPresentation();
            yield this.updateRelation();
            yield this.cleanSlide(this.targetPath);
            yield this.cleanRelations(this.targetRelsPath);
            yield this.checkIntegrity(true, true);
        });
    }
    /**
     * Copys slide layout files
     * @internal
     */
    copySlideLayoutFiles() {
        return __awaiter(this, void 0, void 0, function* () {
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/slideLayouts/slideLayout${this.sourceNumber}.xml`, this.targetArchive, `ppt/slideLayouts/slideLayout${this.targetNumber}.xml`);
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/slideLayouts/_rels/slideLayout${this.sourceNumber}.xml.rels`, this.targetArchive, `ppt/slideLayouts/_rels/slideLayout${this.targetNumber}.xml.rels`);
        });
    }
    updateRelation() {
        return __awaiter(this, void 0, void 0, function* () {
            const layoutToMaster = (yield new xml_relationship_helper_1.XmlRelationshipHelper().initialize(this.targetArchive, `slideLayout${this.targetNumber}.xml.rels`, `ppt/slideLayouts/_rels`, '../slideMasters/slideMaster'));
            layoutToMaster[0].updateTargetIndex(this.targetMaster);
        });
    }
    getName() {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const slideLayoutXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.sourceArchive, `ppt/slideLayouts/slideLayout${this.sourceNumber}.xml`);
            const layout = (_a = slideLayoutXml.getElementsByTagName('p:cSld')) === null || _a === void 0 ? void 0 : _a.item(0);
            if (layout) {
                const name = layout.getAttribute('name');
                return name;
            }
        });
    }
}
exports.Layout = Layout;
//# sourceMappingURL=layout.js.map