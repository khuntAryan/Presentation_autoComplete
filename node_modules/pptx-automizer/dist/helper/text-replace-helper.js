"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const general_helper_1 = require("./general-helper");
const regexp_escape_1 = __importDefault(require("regexp.escape"));
const xml_helper_1 = require("./xml-helper");
const modify_text_helper_1 = __importDefault(require("./modify-text-helper"));
class TextReplaceHelper {
    constructor(options, element) {
        const defaultOptions = {
            openingTag: '{{',
            closingTag: '}}',
        };
        this.options = !options
            ? defaultOptions
            : Object.assign(Object.assign({}, defaultOptions), options);
        this.element = element;
        this.expressions = {
            openingTag: (0, regexp_escape_1.default)(this.options.openingTag),
            closingTag: (0, regexp_escape_1.default)(this.options.closingTag),
        };
    }
    isolateTaggedNodes() {
        const paragraphs = this.element.getElementsByTagName('a:p');
        const pattern = this.getRegExp();
        for (let p = 0; p < paragraphs.length; p++) {
            const blocks = paragraphs[p].getElementsByTagName('a:r');
            for (let r = 0; r < blocks.length; r++) {
                const block = blocks[r];
                const textContent = this.getTextElement(block).textContent;
                const match = textContent.matchAll(pattern);
                const matches = [...match];
                if (matches.length) {
                    this.splitTextBlock(block, matches, textContent);
                }
            }
        }
        // XmlHelper.dump(this.element)
        return this;
    }
    splitTextBlock(block, matches, textContent) {
        const split = this.getCharacterSplit(matches, textContent);
        let lastBlock = block;
        split.forEach((split) => {
            lastBlock = this.insertBlock(lastBlock, split.text);
        });
        block.parentNode.removeChild(block);
    }
    getCharacterSplit(matches, textContent) {
        let lastEnd;
        const split = [];
        matches.forEach((match, s) => {
            const start = match.index;
            const end = match.index + match[0].length;
            if (s === 0 && start > 0) {
                this.pushCharacterSplit(split, 0, start, textContent);
            }
            if (start > lastEnd) {
                this.pushCharacterSplit(split, lastEnd, match.index, textContent);
            }
            this.pushCharacterSplit(split, start, end, textContent);
            const length = textContent.length;
            if (!matches[s + 1] && end < length) {
                this.pushCharacterSplit(split, end, length, textContent);
            }
            lastEnd = end;
        });
        return split;
    }
    pushCharacterSplit(split, from, to, text) {
        split.push({
            from: from,
            to: to,
            text: text.slice(from, to),
        });
    }
    insertBlock(block, text) {
        const newBlock = block.cloneNode(true);
        const newTextElement = this.getTextElement(newBlock);
        modify_text_helper_1.default.content(text)(newTextElement);
        xml_helper_1.XmlHelper.insertAfter(newBlock, block);
        return newBlock;
    }
    applyReplacements(replaceTexts) {
        const textBlocks = this.element.getElementsByTagName('a:r');
        const length = textBlocks.length;
        for (let i = 0; i < length; i++) {
            const textBlock = textBlocks[i];
            replaceTexts.forEach((item) => {
                this.applyReplacement(item, textBlock, i);
            });
        }
    }
    applyReplacement(replaceText, textBlock, currentIndex) {
        var _a;
        const replace = this.options.openingTag + replaceText.replace + this.options.closingTag;
        let textNode = this.getTextElement(textBlock);
        const sourceText = (_a = textNode.firstChild) === null || _a === void 0 ? void 0 : _a.textContent;
        if (sourceText === null || sourceText === void 0 ? void 0 : sourceText.includes(replace)) {
            const bys = general_helper_1.GeneralHelper.arrayify(replaceText.by);
            const modifyBlocks = this.assertTextBlocks(bys.length, textBlock);
            bys.forEach((by, blockIndex) => {
                const textNode = modifyBlocks[blockIndex].getElementsByTagName('a:t')[0];
                this.updateTextNode(textNode, sourceText, replace, by);
            });
        }
    }
    assertTextBlocks(length, textBlock) {
        const modifyBlocks = [];
        if (length > 1) {
            for (let i = 1; i < length; i++) {
                const addedTextBlock = textBlock.cloneNode(true);
                xml_helper_1.XmlHelper.insertAfter(addedTextBlock, textBlock);
                modifyBlocks.push(addedTextBlock);
            }
        }
        modifyBlocks.push(textBlock);
        modifyBlocks.reverse();
        return modifyBlocks;
    }
    updateTextNode(textNode, sourceText, replace, by) {
        const replacedText = sourceText.replace(replace, by.text);
        modify_text_helper_1.default.content(replacedText)(textNode);
        if (by.style) {
            const styleParent = textNode.parentNode;
            const styleElement = styleParent.getElementsByTagName('a:rPr')[0];
            modify_text_helper_1.default.style(by.style)(styleElement);
        }
    }
    getTextElement(block) {
        return block.getElementsByTagName('a:t')[0];
    }
    getRegExp() {
        return new RegExp([
            this.expressions.openingTag,
            '[^',
            this.expressions.openingTag,
            this.expressions.closingTag,
            ']+',
            this.expressions.closingTag,
        ].join(''), 'g');
    }
}
exports.default = TextReplaceHelper;
//# sourceMappingURL=text-replace-helper.js.map