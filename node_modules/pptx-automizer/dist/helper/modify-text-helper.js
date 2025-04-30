"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const modify_color_helper_1 = __importDefault(require("./modify-color-helper"));
const modify_xml_helper_1 = __importDefault(require("./modify-xml-helper"));
const xml_elements_1 = __importDefault(require("./xml-elements"));
class ModifyTextHelper {
}
exports.default = ModifyTextHelper;
/**
 * Set text content of first paragraph and remove remaining block/paragraph elements.
 */
ModifyTextHelper.setText = (text) => (element) => {
    const paragraphs = element.getElementsByTagName('a:p');
    const length = paragraphs.length;
    for (let i = 0; i < length; i++) {
        const paragraph = paragraphs[i];
        if (i === 0) {
            const blocks = element.getElementsByTagName('a:r');
            const length = blocks.length;
            for (let j = 0; j < length; j++) {
                const block = blocks[j];
                if (j === 0) {
                    const textNode = block.getElementsByTagName('a:t')[0];
                    ModifyTextHelper.content(text)(textNode);
                }
                else {
                    block.parentNode.removeChild(block);
                }
            }
        }
        else {
            paragraph.parentNode.removeChild(paragraph);
        }
    }
};
ModifyTextHelper.setBulletList = (list) => (element) => {
    const xmlElements = new xml_elements_1.default(element);
    xmlElements.addBulletList(list);
};
ModifyTextHelper.content = (label) => (element) => {
    if (label !== undefined && element.firstChild) {
        element.firstChild.textContent = String(label);
    }
};
/**
 * Set text style inside an <a:rPr> element
 */
ModifyTextHelper.style = (style) => (element) => {
    if (!style)
        return;
    if (style.color !== undefined) {
        ModifyTextHelper.setColor(style.color)(element);
    }
    if (style.size !== undefined) {
        ModifyTextHelper.setSize(style.size)(element);
    }
    if (style.isBold !== undefined) {
        ModifyTextHelper.setBold(style.isBold)(element);
    }
    if (style.isItalics !== undefined) {
        ModifyTextHelper.setItalics(style.isItalics)(element);
    }
};
/**
 * Set color of text insinde an <a:rPr> element
 */
ModifyTextHelper.setColor = (color) => (element) => {
    modify_color_helper_1.default.solidFill(color)(element);
};
/**
 * Set size of text inside an <a:rPr> element
 */
ModifyTextHelper.setSize = (size) => (element) => {
    if (!size)
        return;
    element.setAttribute('sz', String(Math.round(size)));
};
/**
 * Set bold attribute on text
 */
ModifyTextHelper.setBold = (isBold) => (element) => {
    modify_xml_helper_1.default.booleanAttribute('b', isBold)(element);
};
/**
 * Set italics attribute on text
 */
ModifyTextHelper.setItalics = (isItalics) => (element) => {
    modify_xml_helper_1.default.booleanAttribute('i', isItalics)(element);
};
//# sourceMappingURL=modify-text-helper.js.map