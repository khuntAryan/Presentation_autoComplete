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
const crypto_1 = require("crypto");
const pptxgenjs_1 = __importDefault(require("pptxgenjs"));
/**
 * Using pptxGenJs on an automizer ISlide will create a temporary pptx template
 * and auto-import the generated shapes to the right place on the output slides.
 */
class GeneratePptxGenJs {
    constructor(automizer, slides) {
        this.countSlides = 0;
        /**
         * This is a wrapper around supported pptxGenJS slide item types.
         * It is required to create a unique objectName and find the generated
         * shapes by object name later.
         *
         * @param pgenSlide
         * @param generateElement
         * @param addedObjects
         */
        this.addSlideItems = (pgenSlide, generateElement, addedObjects) => {
            const getObjectName = () => {
                return this.generateObjectName(generateElement, addedObjects);
            };
            return {
                addChart: (type, data, options) => {
                    pgenSlide.addChart(type, data, this.getOptions(options, getObjectName()));
                },
                addImage: (options) => {
                    pgenSlide.addImage(this.getOptions(options, getObjectName()));
                },
                addShape: (shapeName, options) => {
                    pgenSlide.addShape(shapeName, this.getOptions(options, getObjectName()));
                },
                addTable: (tableRows, options) => {
                    pgenSlide.addTable(tableRows, this.getOptions(options, getObjectName()));
                },
                addText: (text, options) => {
                    pgenSlide.addText(text, this.getOptions(options, getObjectName()));
                },
            };
        };
        this.getOptions = (options, objectName) => {
            options = options || {};
            return Object.assign(Object.assign({}, options), { objectName });
        };
        this.automizer = automizer;
        this.slides = slides;
        this.create();
    }
    create() {
        if (this.automizer.params.pptxGenJs) {
            // Use a customized pptxGenJs instance
            this.generator = this.automizer.params.pptxGenJs;
        }
        else {
            // Or the installed version
            this.generator = new pptxgenjs_1.default();
        }
    }
    generateSlides() {
        return __awaiter(this, void 0, void 0, function* () {
            this.tmpFile = (0, crypto_1.randomUUID)() + '.pptx';
            for (const slide of this.slides) {
                const generate = slide.getGeneratedElements();
                if (generate.length) {
                    this.countSlides++;
                    const pgenSlide = this.appendPptxGenSlide();
                    yield this.generateElements(generate, pgenSlide, this.countSlides);
                }
            }
            for (const slide of this.slides) {
                const generate = slide.getGeneratedElements();
                if (generate.length) {
                    this.addElements(generate, slide);
                }
            }
            if (this.countSlides > 0) {
                const data = (yield this.generator.stream());
                this.automizer.load(data, this.tmpFile);
                // await this.generator.writeFile({
                //   fileName: this.automizer.templateDir + '/' + this.tmpFile,
                // });
            }
        });
    }
    generateElements(generate, pgenSlide, tmpSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            for (const generateElement of generate) {
                generateElement.tmpSlideNumber = tmpSlideNumber;
                const addedObjects = [];
                yield generateElement.callback(this.addSlideItems(pgenSlide, generateElement, addedObjects), this.generator);
                generateElement.addedObjects = [...addedObjects];
            }
        });
    }
    addElements(generate, slide) {
        generate.forEach((generateElement) => {
            generateElement.addedObjects.forEach((addedObjectName) => {
                slide.addElement(this.tmpFile, generateElement.tmpSlideNumber, addedObjectName);
            });
        });
    }
    generateObjectName(generateElement, addedObjects) {
        const objectName = (generateElement.objectName ? generateElement.objectName + '-' : '') +
            (0, crypto_1.randomUUID)();
        addedObjects.push(objectName);
        return objectName;
    }
    appendPptxGenSlide() {
        return this.generator.addSlide();
    }
    cleanup() {
        return __awaiter(this, void 0, void 0, function* () {
            // if (this.countSlides > 0) {
            //   fs.unlinkSync(this.automizer.templateDir + '/' + this.tmpFile);
            // }
        });
    }
}
exports.default = GeneratePptxGenJs;
//# sourceMappingURL=generate-pptxgenjs.js.map