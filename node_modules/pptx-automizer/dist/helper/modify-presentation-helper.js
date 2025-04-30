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
var _a;
Object.defineProperty(exports, "__esModule", { value: true });
const xml_helper_1 = require("./xml-helper");
const content_tracker_1 = require("./content-tracker");
const file_helper_1 = require("./file-helper");
class ModifyPresentationHelper {
    /**
     * Tracker.files includes all files that have been
     * copied to the root template by automizer. We remove all other files.
     */
    static removeUnusedFiles(xml, i, archive) {
        return __awaiter(this, void 0, void 0, function* () {
            // Need to skip some dirs until masters and layouts are handled properly
            const skipDirs = [
                'ppt/slideMasters',
                'ppt/slideMasters/_rels',
                'ppt/slideLayouts',
                'ppt/slideLayouts/_rels',
            ];
            for (const dir in content_tracker_1.contentTracker.files) {
                if (skipDirs.includes(dir)) {
                    continue;
                }
                const requiredFiles = content_tracker_1.contentTracker.files[dir];
                yield file_helper_1.FileHelper.removeFromDirectory(archive, dir, (file) => {
                    return !requiredFiles.includes(file.relativePath);
                });
            }
        });
    }
    /**
     * PPT won't complain about unused items in [Content_Types].xml,
     * but we remove them anyway in case the file mentioned in PartName-
     * attribute does not exist.
     */
    static removeUnusedContentTypes(xml, i, archive) {
        return __awaiter(this, void 0, void 0, function* () {
            yield xml_helper_1.XmlHelper.removeIf({
                archive,
                file: `[Content_Types].xml`,
                tag: 'Override',
                clause: (xml, element) => {
                    const filename = element.getAttribute('PartName').substring(1);
                    const exists = file_helper_1.FileHelper.fileExistsInArchive(archive, filename);
                    return exists ? false : true;
                },
            });
        });
    }
    static removedUnusedImages(xml, i, archive) {
        return __awaiter(this, void 0, void 0, function* () {
            yield content_tracker_1.contentTracker.analyzeContents(archive);
            const extensions = ['jpg', 'jpeg', 'png', 'gif', 'svg', 'emf'];
            const keepFiles = [];
            yield content_tracker_1.contentTracker.collect('ppt/slides', 'image', keepFiles);
            yield content_tracker_1.contentTracker.collect('ppt/slideMasters', 'image', keepFiles);
            yield content_tracker_1.contentTracker.collect('ppt/slideLayouts', 'image', keepFiles);
            yield file_helper_1.FileHelper.removeFromDirectory(archive, 'ppt/media', (file) => {
                const info = file_helper_1.FileHelper.getFileInfo(file.name);
                return (extensions.includes(info.extension.toLowerCase()) &&
                    !keepFiles.includes(info.base));
            });
        });
    }
}
exports.default = ModifyPresentationHelper;
_a = ModifyPresentationHelper;
/**
 * Get Collection of slides
 */
ModifyPresentationHelper.getSlidesCollection = (xml) => {
    return xml.getElementsByTagName('p:sldId');
};
ModifyPresentationHelper.getSlideMastersCollection = (xml) => {
    return xml.getElementsByTagName('p:sldMasterId');
};
/**
 * Pass an array of slide numbers to define a target sort order.
 * First slide starts by 1.
 * @order Array of slide numbers, starting by 1
 */
ModifyPresentationHelper.sortSlides = (order) => (xml) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    order.map((index, i) => order[i]--);
    xml_helper_1.XmlHelper.sortCollection(slides, order);
};
/**
 * Set ids to prevent corrupted pptx.
 * Must start with 256 and increment by one.
 */
ModifyPresentationHelper.normalizeSlideIds = (xml) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    const firstId = 256;
    xml_helper_1.XmlHelper.modifyCollection(slides, (slide, i) => {
        slide.setAttribute('id', String(firstId + i));
    });
};
/**
 * Update slideMaster ids to prevent corrupted pptx.
 * - Take first slideMaster id from presentation.xml to start,
 * - then update incremental ids of each p:sldLayoutId in slideMaster[i].xml
 *   (starting by slideMasterId + 1)
 * - and update next slideMaster id with previous p:sldLayoutId + 1
 *
 * p:sldMasterId-ids and p:sldLayoutId-ids need to be in a row, otherwise
 * PowerPoint will complain on any p:sldLayoutId-id lower than its
 * corresponding slideMaster-id. omg.
 */
ModifyPresentationHelper.normalizeSlideMasterIds = (xml, i, archive) => __awaiter(void 0, void 0, void 0, function* () {
    const slides = ModifyPresentationHelper.getSlideMastersCollection(xml);
    let currentId;
    yield xml_helper_1.XmlHelper.modifyCollectionAsync(slides, (slide, i) => __awaiter(void 0, void 0, void 0, function* () {
        const masterId = i + 1;
        if (i === 0) {
            currentId = Number(slide.getAttribute('id'));
        }
        slide.setAttribute('id', String(currentId));
        currentId++;
        const slideMasterXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, `ppt/slideMasters/slideMaster${masterId}.xml`);
        const slideLayouts = slideMasterXml.getElementsByTagName('p:sldLayoutId');
        xml_helper_1.XmlHelper.modifyCollection(slideLayouts, (slideLayout) => {
            slideLayout.setAttribute('id', String(currentId));
            currentId++;
        });
    }));
});
//# sourceMappingURL=modify-presentation-helper.js.map