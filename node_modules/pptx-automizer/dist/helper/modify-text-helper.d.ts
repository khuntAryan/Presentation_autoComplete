import { Color, TextStyle } from '../types/modify-types';
import { XmlElement } from '../types/xml-types';
export default class ModifyTextHelper {
    /**
     * Set text content of first paragraph and remove remaining block/paragraph elements.
     */
    static setText: (text: number | string) => (element: XmlElement) => void;
    static setBulletList: (list: any) => (element: XmlElement) => void;
    static content: (label: number | string | undefined) => (element: XmlElement) => void;
    /**
     * Set text style inside an <a:rPr> element
     */
    static style: (style: TextStyle) => (element: XmlElement) => void;
    /**
     * Set color of text insinde an <a:rPr> element
     */
    static setColor: (color: Color) => (element: XmlElement) => void;
    /**
     * Set size of text inside an <a:rPr> element
     */
    static setSize: (size: number) => (element: XmlElement) => void;
    /**
     * Set bold attribute on text
     */
    static setBold: (isBold: boolean) => (element: XmlElement) => void;
    /**
     * Set italics attribute on text
     */
    static setItalics: (isItalics: boolean) => (element: XmlElement) => void;
}
