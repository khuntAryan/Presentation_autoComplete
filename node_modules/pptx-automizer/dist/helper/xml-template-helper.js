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
exports.XmlTemplateHelper = void 0;
const xml_helper_1 = require("./xml-helper");
const xml_relationship_helper_1 = require("./xml-relationship-helper");
const xml_slide_helper_1 = require("./xml-slide-helper");
class XmlTemplateHelper {
    constructor(archive) {
        this.relType =
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
        this.relTypeNotes =
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
        this.archive = archive;
        this.path = 'ppt/_rels/presentation.xml.rels';
        this.defaultSlideName = 'untitled';
    }
    getCreationIds() {
        return __awaiter(this, void 0, void 0, function* () {
            const archive = this.archive;
            const relationships = yield xml_helper_1.XmlHelper.getTargetsByRelationshipType(archive, this.path, this.relType);
            const creationIds = [];
            for (const slideRel of relationships) {
                try {
                    const slideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, 'ppt/' + slideRel.file);
                    if (!slideXml) {
                        console.warn(`slideXml is undefined for file ${slideRel.file}`);
                        continue;
                    }
                    const slideHelper = new xml_slide_helper_1.XmlSlideHelper(slideXml);
                    const creationIdSlide = slideHelper.getSlideCreationId();
                    if (!creationIdSlide) {
                        console.warn(`No creationId found in ${slideRel.file}`);
                    }
                    const slideInfo = yield this.getSlideInfo(slideXml, archive, slideRel.file);
                    creationIds.push({
                        id: creationIdSlide,
                        number: this.parseSlideRelFile(slideRel.file),
                        elements: slideHelper.getAllElements(),
                        info: slideInfo,
                    });
                }
                catch (err) {
                    console.error(`An error occurred while processing ${slideRel.file}:`, err);
                }
            }
            return creationIds.sort((slideA, slideB) => slideA.number < slideB.number ? -1 : 1);
        });
    }
    parseSlideRelFile(slideRelFile) {
        return Number(slideRelFile.replace('slides/slide', '').replace('.xml', ''));
    }
    getSlideInfo(slideXml, archive, slideRelFile) {
        return __awaiter(this, void 0, void 0, function* () {
            let name;
            const slideNoteRels = yield this.getSlideNoteRels(archive, slideRelFile);
            if (slideNoteRels.length > 0) {
                name = yield this.getSlideNameFromNotes(archive, slideNoteRels);
            }
            if (!name) {
                name = this.getNameFromSlideInfo(slideXml);
            }
            name = !name ? this.defaultSlideName : name;
            return {
                name: name,
            };
        });
    }
    getNameFromSlideInfo(slideXml) {
        const slideTitle = slideXml.getElementsByTagName('p:ph');
        if (slideTitle.length && slideTitle[0].getAttribute('type') === 'title') {
            const titleElement = slideTitle[0].parentNode.parentNode
                .parentNode;
            const nameFragments = this.parseTitleElement(titleElement);
            if (nameFragments.length) {
                return nameFragments.join(' ');
            }
        }
    }
    getSlideNoteRels(archive, slideRelFile) {
        return __awaiter(this, void 0, void 0, function* () {
            const relFileName = slideRelFile.replace('slides', '');
            const slideRels = yield xml_helper_1.XmlHelper.getTargetsByRelationshipType(archive, `ppt/slides/_rels${relFileName}.rels`, this.relTypeNotes);
            return slideRels;
        });
    }
    getSlideNameFromNotes(archive, slideNoteRels) {
        return __awaiter(this, void 0, void 0, function* () {
            const notesFile = slideNoteRels[0].file.replace('../', '');
            const notesXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, 'ppt/' + notesFile);
            const titleElements = notesXml.getElementsByTagName('a:p');
            if (titleElements.length > 0) {
                const nameFragments = this.parseTitleElement(titleElements[0]);
                if (nameFragments.length) {
                    return nameFragments.join('');
                }
            }
        });
    }
    parseTitleElement(titleElement) {
        var _a;
        const nameFragments = [];
        const titleText = titleElement.getElementsByTagName('a:t');
        if (titleText.length) {
            for (const titleTextNode in titleText) {
                if ((_a = titleText[titleTextNode].firstChild) === null || _a === void 0 ? void 0 : _a.nodeValue) {
                    nameFragments.push(titleText[titleTextNode].firstChild.nodeValue);
                }
            }
        }
        return nameFragments;
    }
    /**
     * Returns the slide numbers of a given template as a sorted array of integers.
     * @returns {Promise<number[]>} - A promise that resolves to a sorted array of slide numbers in the template.
     */
    getAllSlideNumbers() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const archive = this.archive;
                const xmlRelationshipHelper = new xml_relationship_helper_1.XmlRelationshipHelper();
                const allSlides = (yield xmlRelationshipHelper.initialize(archive, 'presentation.xml.rels', 'ppt/_rels', 'slides/slide'));
                // Extract slide numbers from each slide using the 'number' property and sort the array of integers.
                const slideNumbers = allSlides.map((slide) => slide.number);
                slideNumbers.sort((a, b) => a - b);
                return slideNumbers;
            }
            catch (error) {
                throw new Error(`Error getting slide numbers: ${error.message}`);
            }
        });
    }
}
exports.XmlTemplateHelper = XmlTemplateHelper;
//# sourceMappingURL=xml-template-helper.js.map