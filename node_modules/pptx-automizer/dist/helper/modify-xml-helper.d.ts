import { Modification, ModificationTags } from '../types/modify-types';
import { XmlDocument, XmlElement } from '../types/xml-types';
export default class ModifyXmlHelper {
    root: XmlDocument | XmlElement;
    templates: {
        [key: string]: XmlElement;
    };
    constructor(root: XmlDocument | XmlElement);
    modify(tags: ModificationTags, root?: XmlDocument | XmlElement): void;
    modifyAll(tag: string, modifier: Modification, root: XmlDocument | XmlElement): void;
    assertElement(collection: HTMLCollectionOf<Element>, index: number, tag: string, parent: XmlDocument | XmlElement, modifier: Modification): XmlDocument | XmlElement | boolean;
    createElement(parent: XmlDocument | XmlElement, tag: string): boolean;
    static getText: (element: XmlElement) => string;
    static value: (value: number | string, index?: number) => (element: XmlElement) => void;
    static textContent: (value: number | string) => (element: XmlElement) => void;
    static attribute: (attribute: string, value: string | number) => (element: XmlElement) => void;
    static booleanAttribute: (attribute: string, state: boolean) => (element: XmlElement) => void;
    static range: (series: number, length?: number) => (element: XmlElement) => void;
}
