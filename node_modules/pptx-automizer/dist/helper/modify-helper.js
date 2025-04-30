"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DxaToCm = exports.CmToDxa = void 0;
const xml_helper_1 = require("./xml-helper");
class ModifyHelper {
}
exports.default = ModifyHelper;
/**
 * Set value of an attribute.
 * @param tagName specify the tag name to search for
 * @param attribute name of target attribute
 * @param value the value to be set on the attribute
 * @param [count] specify if element index is different to zero
 */
ModifyHelper.setAttribute = (tagName, attribute, value, count) => (element) => {
    const item = element.getElementsByTagName(tagName)[count || 0];
    if (item.setAttribute !== undefined) {
        item.setAttribute(attribute, String(value));
    }
};
/**
 * Dump current element to console.
 */
ModifyHelper.dump = (element) => {
    xml_helper_1.XmlHelper.dump(element);
};
/**
 * Dump current chart to console.
 */
ModifyHelper.dumpChart = (element, chart) => {
    xml_helper_1.XmlHelper.dump(chart);
};
/*
  Convert cm to ppt's dxa unit
 */
const CmToDxa = (cm) => {
    return Math.round(cm * 360000);
};
exports.CmToDxa = CmToDxa;
/*
  Convert ppt's dxa unit to cm
 */
const DxaToCm = (dxa) => {
    return dxa / 360000;
};
exports.DxaToCm = DxaToCm;
//# sourceMappingURL=modify-helper.js.map