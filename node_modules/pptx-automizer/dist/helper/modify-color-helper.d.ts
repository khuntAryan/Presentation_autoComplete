import { Color } from '../types/modify-types';
import { XmlElement } from '../types/xml-types';
export default class ModifyColorHelper {
    /**
     * Replaces or creates an <a:solidFill> Element
     */
    static solidFill: (color: Color, index?: number | 'last') => (element: XmlElement) => void;
    static removeNoFill: () => (element: XmlElement) => void;
    static normalizeColorObject: (color: Color) => Color;
}
