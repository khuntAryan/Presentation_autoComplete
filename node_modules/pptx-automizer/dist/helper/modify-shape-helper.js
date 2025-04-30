"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const general_helper_1 = require("./general-helper");
const text_replace_helper_1 = __importDefault(require("./text-replace-helper"));
const modify_text_helper_1 = __importDefault(require("./modify-text-helper"));
const map = {
    x: { tag: 'a:off', attribute: 'x' },
    l: { tag: 'a:off', attribute: 'x' },
    left: { tag: 'a:off', attribute: 'x' },
    y: { tag: 'a:off', attribute: 'y' },
    t: { tag: 'a:off', attribute: 'y' },
    top: { tag: 'a:off', attribute: 'y' },
    cx: { tag: 'a:ext', attribute: 'cx' },
    w: { tag: 'a:ext', attribute: 'cx' },
    width: { tag: 'a:ext', attribute: 'cx' },
    cy: { tag: 'a:ext', attribute: 'cy' },
    h: { tag: 'a:ext', attribute: 'cy' },
    height: { tag: 'a:ext', attribute: 'cy' },
};
class ModifyShapeHelper {
}
exports.default = ModifyShapeHelper;
/**
 * Set solid fill of modified shape
 */
ModifyShapeHelper.setSolidFill = (element) => {
    element
        .getElementsByTagName('a:solidFill')[0]
        .getElementsByTagName('a:schemeClr')[0]
        .setAttribute('val', 'accent6');
};
/**
 * Set text content of modified shape
 */
ModifyShapeHelper.setText = (text) => (element) => {
    modify_text_helper_1.default.setText(text)(element);
};
/**
 * Set content to bulleted list of modified shape
 */
ModifyShapeHelper.setBulletList = (list) => (element) => {
    modify_text_helper_1.default.setBulletList(list)(element);
};
/**
 * Replace tagged text content within modified shape
 */
ModifyShapeHelper.replaceText = (replaceText, options) => (element) => {
    const replaceTexts = general_helper_1.GeneralHelper.arrayify(replaceText);
    new text_replace_helper_1.default(options, element)
        .isolateTaggedNodes()
        .applyReplacements(replaceTexts);
};
/**
 * Set position and size of modified shape.
 */
ModifyShapeHelper.setPosition = (pos) => (element) => {
    const aOff = element.getElementsByTagName('a:off');
    if (!(aOff === null || aOff === void 0 ? void 0 : aOff.item(0))) {
        return;
    }
    const xfrm = aOff.item(0).parentNode;
    Object.keys(pos).forEach((key) => {
        let value = Math.round(pos[key]);
        if (typeof value !== 'number' || !map[key])
            return;
        value = value < 0 ? 0 : value;
        xfrm
            .getElementsByTagName(map[key].tag)[0]
            .setAttribute(map[key].attribute, value);
    });
};
/**
 * Update position and size of a shape by a given Value.
 */
ModifyShapeHelper.updatePosition = (pos) => (element) => {
    const xfrm = element.getElementsByTagName('a:off')[0]
        .parentNode;
    Object.keys(pos).forEach((key) => {
        let value = Math.round(pos[key]);
        if (typeof value !== 'number' || !map[key])
            return;
        const currentValue = xfrm
            .getElementsByTagName(map[key].tag)[0]
            .getAttribute(map[key].attribute);
        value += Number(currentValue);
        xfrm
            .getElementsByTagName(map[key].tag)[0]
            .setAttribute(map[key].attribute, value);
    });
};
/**
 * Rotate a shape by a given value. Use e.g. 180 to flip a shape.
 * A negative value will rotate counter clockwise.
 * @param degrees Rotate by °
 */
ModifyShapeHelper.rotate = (degrees) => (element) => {
    const spPr = element.getElementsByTagName('p:spPr');
    if (spPr) {
        const xfrm = spPr.item(0).getElementsByTagName('a:xfrm').item(0);
        degrees = degrees < 0 ? 360 + degrees : degrees;
        xfrm.setAttribute('rot', String(Math.round(degrees * 60000)));
    }
};
//# sourceMappingURL=modify-shape-helper.js.map