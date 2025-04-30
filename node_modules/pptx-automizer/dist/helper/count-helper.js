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
exports.CountHelper = void 0;
const xml_helper_1 = require("./xml-helper");
class CountHelper {
    constructor(name, template) {
        this.name = name;
        this.template = template;
    }
    static increment(name, counters) {
        return CountHelper.getCounterByName(name, counters)._increment();
    }
    static count(name, counters) {
        return CountHelper.getCounterByName(name, counters).get();
    }
    static reset(counters) {
        counters.forEach((counter) => (counter.count = undefined));
    }
    static getCounterByName(name, counters) {
        const counter = counters.find((c) => c.name === name);
        if (counter === undefined) {
            throw new Error(`Counter ${name} not found.`);
        }
        return counter;
    }
    _increment() {
        this.count++;
        return this.count;
    }
    set() {
        return __awaiter(this, void 0, void 0, function* () {
            this.count = yield this.calculateCount(yield this.template.archive);
        });
    }
    get() {
        return this.count;
    }
    calculateCount(presentation) {
        switch (this.name) {
            case 'slides':
                return CountHelper.countSlides(presentation);
            case 'masters':
                return CountHelper.countMasters(presentation);
            case 'layouts':
                return CountHelper.countLayouts(presentation);
            case 'themes':
                return CountHelper.countThemes(presentation);
            case 'charts':
                return CountHelper.countCharts(presentation);
            case 'images':
                return CountHelper.countImages(presentation);
            case 'oleObjects':
                return CountHelper.countOleObjects(presentation);
        }
        throw new Error(`No way to count ${this.name}.`);
    }
    static countSlides(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const presentationXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(presentation, 'ppt/presentation.xml');
            return presentationXml.getElementsByTagName('p:sldId').length;
        });
    }
    static countMasters(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const presentationXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(presentation, 'ppt/presentation.xml');
            return presentationXml.getElementsByTagName('p:sldMasterId').length;
        });
    }
    static countLayouts(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const contentTypesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(presentation, '[Content_Types].xml');
            const overrides = contentTypesXml.getElementsByTagName('Override');
            return Object.keys(overrides)
                .map((key) => overrides[key])
                .filter((o) => o.getAttribute &&
                o.getAttribute('ContentType') ===
                    `application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml`).length;
        });
    }
    static countThemes(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const contentTypesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(presentation, '[Content_Types].xml');
            const overrides = contentTypesXml.getElementsByTagName('Override');
            return Object.keys(overrides)
                .map((key) => overrides[key])
                .filter((o) => o.getAttribute &&
                o.getAttribute('ContentType') ===
                    `application/vnd.openxmlformats-officedocument.theme+xml`).length;
        });
    }
    static countCharts(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const contentTypesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(presentation, '[Content_Types].xml');
            const overrides = contentTypesXml.getElementsByTagName('Override');
            return Object.keys(overrides)
                .map((key) => overrides[key])
                .filter((o) => o.getAttribute &&
                o.getAttribute('ContentType') ===
                    `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`).length;
        });
    }
    static countOleObjects(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const contentTypesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(presentation, '[Content_Types].xml');
            const overrides = contentTypesXml.getElementsByTagName('Override');
            return Object.keys(overrides)
                .map((key) => overrides[key])
                .filter((o) => o.getAttribute &&
                o.getAttribute('ContentType') ===
                    `application/vnd.openxmlformats-officedocument.oleObject`).length;
        });
    }
    static countImages(presentation) {
        return __awaiter(this, void 0, void 0, function* () {
            const mediaFiles = yield presentation.folder('ppt/media');
            const count = mediaFiles.filter((file) => file.relativePath.indexOf('image') === 0).length;
            return count;
        });
    }
}
exports.CountHelper = CountHelper;
//# sourceMappingURL=count-helper.js.map