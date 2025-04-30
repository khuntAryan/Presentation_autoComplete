"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const modify_color_helper_1 = __importDefault(require("./modify-color-helper"));
const xml_helper_1 = require("./xml-helper");
const index_1 = require("../index");
class ModifyBackgroundHelper {
}
exports.default = ModifyBackgroundHelper;
/**
 * Set solid fill of master background
 */
ModifyBackgroundHelper.setSolidFill = (color) => (slideMasterXml) => {
    var _a;
    const bgPr = (_a = slideMasterXml.getElementsByTagName('p:bgPr')) === null || _a === void 0 ? void 0 : _a.item(0);
    if (bgPr) {
        modify_color_helper_1.default.solidFill(color)(bgPr);
    }
    else {
        throw 'No background properties for slideMaster';
    }
};
/**
 * Modify a slideMaster background image's relation target
 * @param master
 * @param imageName
 */
ModifyBackgroundHelper.setRelationTarget = (master, imageName) => {
    let targetRelation = '';
    master.modify((masterXml) => {
        targetRelation =
            ModifyBackgroundHelper.getBackgroundProperties(masterXml);
    });
    master.modifyRelations((relXml) => {
        const relations = xml_helper_1.XmlHelper.findByAttributeValue(relXml.getElementsByTagName('Relationship'), 'Id', targetRelation);
        if (relations[0]) {
            index_1.ModifyImageHelper.setRelationTarget(imageName)(undefined, relations[0]);
        }
    });
};
/**
 * Extract background properties from slideMaster xml
 */
ModifyBackgroundHelper.getBackgroundProperties = (slideMasterXml) => {
    var _a, _b;
    const bgPr = (_a = slideMasterXml.getElementsByTagName('p:bgPr')) === null || _a === void 0 ? void 0 : _a.item(0);
    if (bgPr) {
        const blip = (_b = bgPr
            .getElementsByTagName('a:blip')) === null || _b === void 0 ? void 0 : _b.item(0).getAttribute('r:embed');
        return blip;
    }
    else {
        throw 'No background properties for slideMaster';
    }
};
//# sourceMappingURL=modify-background-helper.js.map