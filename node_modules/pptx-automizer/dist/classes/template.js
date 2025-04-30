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
exports.Template = void 0;
const file_helper_1 = require("../helper/file-helper");
const count_helper_1 = require("../helper/count-helper");
const xml_template_helper_1 = require("../helper/xml-template-helper");
const xml_helper_1 = require("../helper/xml-helper");
const generate_pptxgenjs_1 = __importDefault(require("../helper/generate/generate-pptxgenjs"));
class Template {
    constructor(file, params) {
        this.contentMap = [];
        this.mediaFiles = [];
        this.file = file;
        const archive = file_helper_1.FileHelper.importArchive(file, params);
        this.archive = archive;
    }
    static import(file, params, automizer) {
        let newTemplate;
        if (params.name) {
            // New template will be a default template containing
            // importable slides and shapes.
            newTemplate = new Template(file, params);
            newTemplate.name = params.name;
        }
        else {
            // New template will be root template
            newTemplate = new Template(file, params);
            newTemplate.automizer = automizer;
            newTemplate.slides = [];
            newTemplate.masters = [];
            newTemplate.counter = [
                new count_helper_1.CountHelper('slides', newTemplate),
                new count_helper_1.CountHelper('charts', newTemplate),
                new count_helper_1.CountHelper('images', newTemplate),
                new count_helper_1.CountHelper('masters', newTemplate),
                new count_helper_1.CountHelper('layouts', newTemplate),
                new count_helper_1.CountHelper('themes', newTemplate),
                new count_helper_1.CountHelper('oleObjects', newTemplate),
            ];
            // TODO: refactor content tracker, let root template have an instance
            // newTemplate.content = new ContentTracker();
        }
        return newTemplate;
    }
    mapContents(type, key, sourceId, targetId, name) {
        this.contentMap.push({
            type,
            key,
            sourceId,
            targetId,
            name,
        });
    }
    getNamedMappedContent(type, name) {
        return this.contentMap.find((map) => map.type === type && map.name === name);
    }
    getMappedContent(type, key, sourceId) {
        return this.contentMap.find((map) => map.type === type && map.key === key && map.sourceId === sourceId);
    }
    /**
     * Returns the slide numbers of a given template as a sorted array of integers.
     * @returns {Promise<number[]>} - A promise that resolves to a sorted array of slide numbers in the template.
     */
    getAllSlideNumbers() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const xmlTemplateHelper = new xml_template_helper_1.XmlTemplateHelper(this.archive);
                this.slideNumbers = yield xmlTemplateHelper.getAllSlideNumbers();
                return this.slideNumbers;
            }
            catch (error) {
                throw new Error(error.message);
            }
        });
    }
    setCreationIds() {
        return __awaiter(this, void 0, void 0, function* () {
            const archive = yield this.archive;
            const xmlTemplateHelper = new xml_template_helper_1.XmlTemplateHelper(archive);
            this.creationIds = yield xmlTemplateHelper.getCreationIds();
            return this.creationIds;
        });
    }
    appendMasterSlide(slideMaster) {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.counter[0].get() === undefined) {
                yield this.initializeCounter();
            }
            yield slideMaster.append(this).catch((e) => {
                throw e;
            });
        });
    }
    appendSlide(slide) {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.counter[0].get() === undefined) {
                yield this.initializeCounter();
            }
            yield slide.append(this).catch((e) => {
                throw e;
            });
        });
    }
    appendLayout(slideLayout) {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.counter[0].get() === undefined) {
                yield this.initializeCounter();
            }
            yield slideLayout.append(this).catch((e) => {
                throw e;
            });
        });
    }
    countExistingSlides() {
        return __awaiter(this, void 0, void 0, function* () {
            const xml = yield this.getSlideIdList();
            const sldIdLst = xml.getElementsByTagName('p:sldIdLst');
            if (sldIdLst.length > 0) {
                const existingSlides = sldIdLst[0].getElementsByTagName('p:sldId');
                this.existingSlides = existingSlides.length;
            }
        });
    }
    truncate() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.existingSlides > 0) {
                const xml = yield this.getSlideIdList();
                const existingSlides = xml.getElementsByTagName('p:sldId');
                xml_helper_1.XmlHelper.sliceCollection(existingSlides, this.existingSlides, 0);
                xml_helper_1.XmlHelper.writeXmlToArchive(yield this.archive, `ppt/presentation.xml`, xml);
            }
        });
    }
    getSlideIdList() {
        return __awaiter(this, void 0, void 0, function* () {
            const archive = yield this.archive;
            const xml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, `ppt/presentation.xml`);
            return xml;
        });
    }
    initializeCounter() {
        return __awaiter(this, void 0, void 0, function* () {
            for (const c of this.counter) {
                yield c.set();
            }
        });
    }
    incrementCounter(name) {
        return count_helper_1.CountHelper.increment(name, this.counter);
    }
    count(name) {
        return count_helper_1.CountHelper.count(name, this.counter);
    }
    runExternalGenerator() {
        return __awaiter(this, void 0, void 0, function* () {
            this.generator = new generate_pptxgenjs_1.default(this.automizer, this.slides);
            yield this.generator.generateSlides();
        });
    }
    cleanupExternalGenerator() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.generator.cleanup();
        });
    }
}
exports.Template = Template;
//# sourceMappingURL=template.js.map