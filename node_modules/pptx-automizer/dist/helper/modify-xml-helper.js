"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const cell_id_helper_1 = __importDefault(require("./cell-id-helper"));
const general_helper_1 = require("./general-helper");
const xml_helper_1 = require("./xml-helper");
const xml_elements_1 = __importDefault(require("./xml-elements"));
class ModifyXmlHelper {
    constructor(root) {
        this.root = root;
        this.templates = {};
    }
    modify(tags, root) {
        root = root || this.root;
        for (const tag in tags) {
            const modifier = tags[tag];
            if (modifier.all) {
                this.modifyAll(tag, modifier, root);
            }
            if (modifier.collection) {
                const modifies = general_helper_1.GeneralHelper.arrayify(modifier.collection);
                const collection = root.getElementsByTagName(tag);
                Object.values(modifies).forEach((modifyXml) => modifyXml(collection));
                return;
            }
            const index = modifier.index || 0;
            const isRequired = modifier.isRequired !== undefined ? modifier.isRequired : true;
            const element = this.assertElement(root.getElementsByTagName(tag), index, tag, root, modifier);
            if (element === false) {
                if (isRequired === true) {
                    // vd('Could not assert required tag: ' + tag + '@index:' + index);
                }
            }
            else {
                if (modifier.modify) {
                    const modifies = general_helper_1.GeneralHelper.arrayify(modifier.modify);
                    Object.values(modifies).forEach((modifyXml) => modifyXml(element));
                }
                if (modifier.children) {
                    this.modify(modifier.children, element);
                }
            }
        }
    }
    modifyAll(tag, modifier, root) {
        const elements = Array.from(root.getElementsByTagName(tag));
        elements.forEach((element) => {
            this.modify(modifier.children, element);
        });
    }
    assertElement(collection, index, tag, parent, modifier) {
        if (!collection[index]) {
            if (collection[collection.length - 1] === undefined) {
                this.createElement(parent, tag);
            }
            else {
                const lastSibling = collection[collection.length - 1];
                let sourceSibling = lastSibling;
                if (modifier.fromIndex && collection.item(modifier.fromIndex)) {
                    sourceSibling = collection.item(modifier.fromIndex);
                }
                else if (modifier.fromPrevious && collection.item(index - 1)) {
                    sourceSibling = collection.item(index - 1);
                }
                if ((!sourceSibling || modifier.forceCreate) && this.templates[tag]) {
                    sourceSibling = this.templates[tag];
                }
                const newChild = sourceSibling.cloneNode(true);
                xml_helper_1.XmlHelper.insertAfter(newChild, lastSibling);
            }
        }
        const element = parent.getElementsByTagName(tag)[index];
        if (element) {
            this.templates[tag] =
                this.templates[tag] || element.cloneNode(true);
            return element;
        }
        return false;
    }
    createElement(parent, tag) {
        switch (tag) {
            case 'a:t':
                new xml_elements_1.default(parent).text();
                return true;
            case 'c:dPt':
                new xml_elements_1.default(parent).dataPoint();
                return true;
            case 'c:spPr':
                new xml_elements_1.default(parent).shapeProperties();
                return true;
            case 'c:dLbl':
                new xml_elements_1.default(parent).dataPointLabel();
                return true;
            case 'a:lnL':
            case 'a:lnR':
            case 'a:lnT':
            case 'a:lnB':
                new xml_elements_1.default(parent).tableCellBorder(tag);
                return true;
        }
        return false;
    }
}
exports.default = ModifyXmlHelper;
ModifyXmlHelper.getText = (element) => {
    return element.firstChild.textContent;
};
ModifyXmlHelper.value = (value, index) => (element) => {
    const valueElement = element.getElementsByTagName('c:v');
    if (!valueElement.length) {
        xml_helper_1.XmlHelper.dump(element);
        throw 'Unable to set value @index: ' + index;
    }
    valueElement[0].firstChild.textContent = String(value);
    if (index !== undefined) {
        element.setAttribute('idx', String(index));
    }
};
ModifyXmlHelper.textContent = (value) => (element) => {
    element.firstChild.textContent = String(value);
};
ModifyXmlHelper.attribute = (attribute, value) => (element) => {
    if (value != undefined)
        element.setAttribute(attribute, String(value));
};
ModifyXmlHelper.booleanAttribute = (attribute, state) => (element) => {
    element.setAttribute(attribute, state === true ? '1' : '0');
};
ModifyXmlHelper.range = (series, length) => (element) => {
    const range = element.firstChild.textContent;
    element.firstChild.textContent = cell_id_helper_1.default.setRange(range, series, length);
};
//# sourceMappingURL=modify-xml-helper.js.map