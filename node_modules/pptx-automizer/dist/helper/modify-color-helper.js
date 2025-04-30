"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const xml_elements_1 = __importDefault(require("./xml-elements"));
const xml_helper_1 = require("./xml-helper");
class ModifyColorHelper {
}
exports.default = ModifyColorHelper;
/**
 * Replaces or creates an <a:solidFill> Element
 */
ModifyColorHelper.solidFill = (color, index) => (element) => {
    if (!color || !color.type || (element === null || element === void 0 ? void 0 : element.getElementsByTagName) === undefined)
        return;
    ModifyColorHelper.normalizeColorObject(color);
    const solidFills = element.getElementsByTagName('a:solidFill');
    if (!solidFills.length) {
        const solidFill = new xml_elements_1.default(element, {
            color: color,
        }).solidFill();
        element.appendChild(solidFill);
        return;
    }
    let targetIndex = !index
        ? 0
        : index === 'last'
            ? solidFills.length - 1
            : index;
    const solidFill = solidFills[targetIndex];
    const colorType = new xml_elements_1.default(element, {
        color: color,
    }).colorType();
    xml_helper_1.XmlHelper.sliceCollection(solidFill.childNodes, 0);
    solidFill.appendChild(colorType);
};
ModifyColorHelper.removeNoFill = () => (element) => {
    const hasNoFill = element.getElementsByTagName('a:noFill')[0];
    if (hasNoFill) {
        element.removeChild(hasNoFill);
    }
};
ModifyColorHelper.normalizeColorObject = (color) => {
    if (color.value.indexOf('#') === 0) {
        color.value = color.value.replace('#', '');
    }
    if (color.value.toLowerCase().indexOf('rgb(') === 0) {
        // TODO: convert RGB to HEX
        color.value = 'cccccc';
    }
    return color;
};
//# sourceMappingURL=modify-color-helper.js.map