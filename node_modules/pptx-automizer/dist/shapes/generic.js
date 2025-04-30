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
exports.GenericShape = void 0;
const shape_1 = require("../classes/shape");
class GenericShape extends shape_1.Shape {
    constructor(shape, targetType) {
        super(shape, targetType);
    }
    modify(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.replaceIntoSlideTree();
            return this;
        });
    }
    append(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.appendToSlideTree();
            return this;
        });
    }
    remove(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.removeFromSlideTree();
            return this;
        });
    }
    prepare(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.setTarget(targetTemplate, targetSlideNumber);
            yield this.setTargetElement();
            this.applyCallbacks(this.callbacks, this.targetElement);
        });
    }
}
exports.GenericShape = GenericShape;
//# sourceMappingURL=generic.js.map