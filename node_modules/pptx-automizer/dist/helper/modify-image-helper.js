"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const slugify_1 = __importDefault(require("slugify"));
class ModifyImageHelper {
}
exports.default = ModifyImageHelper;
/**
 * Update the "Target" attribute of a created image relation.
 * This will change the image itself. Load images with Automizer.loadMedia
 * @param filename name of target image in root template media folder.
 */
ModifyImageHelper.setRelationTarget = (filename) => {
    return (element, arg1) => {
        arg1.setAttribute('Target', '../media/' + (0, slugify_1.default)(filename));
    };
};
/*
  Update an existing duotone image overlay element (WIP)
  Apply a duotone color to an image p:blipFill -> a:blip fill element.
  Works best on white icons, see __tests__/media/feather.png
 */
ModifyImageHelper.setDuotoneFill = (duotoneParams) => (element) => {
    const blipFill = element.getElementsByTagName('p:blipFill');
    if (!blipFill) {
        return;
    }
    const duotone = blipFill.item(0).getElementsByTagName('a:duotone')[0];
    if (duotone) {
        if (duotoneParams === null || duotoneParams === void 0 ? void 0 : duotoneParams.color) {
            const srgbClr = duotone.getElementsByTagName('a:srgbClr')[0];
            if (srgbClr) {
                // Only sRgb supported
                srgbClr.setAttribute('val', String(duotoneParams.color.value));
                if ((duotoneParams === null || duotoneParams === void 0 ? void 0 : duotoneParams.tint) !== undefined) {
                    // tint needs to be 0 - 100000
                    const tint = srgbClr.getElementsByTagName('a:tint')[0];
                    if (tint) {
                        tint.setAttribute('val', String(duotoneParams.tint));
                    }
                }
                if ((duotoneParams === null || duotoneParams === void 0 ? void 0 : duotoneParams.satMod) !== undefined) {
                    const satMod = srgbClr.getElementsByTagName('a:satMod')[0];
                    if (satMod) {
                        satMod.setAttribute('val', String(duotoneParams.satMod));
                    }
                }
            }
        }
        if (duotoneParams === null || duotoneParams === void 0 ? void 0 : duotoneParams.prstClr) {
            const prstClr = duotone.getElementsByTagName('a:prstClr')[0];
            if (prstClr) {
                // Only tested with "black" and "white"
                prstClr.setAttribute('val', String(duotoneParams.prstClr));
            }
        }
    }
};
//# sourceMappingURL=modify-image-helper.js.map