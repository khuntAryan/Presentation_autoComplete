import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import { XmlElement } from '../types/xml-types';
export default class ModifyShapeHelper {
    /**
     * Set solid fill of modified shape
     */
    static setSolidFill: (element: XmlElement) => void;
    /**
     * Set text content of modified shape
     */
    static setText: (text: string) => (element: XmlElement) => void;
    /**
     * Set content to bulleted list of modified shape
     */
    static setBulletList: (list: any) => (element: XmlElement) => void;
    /**
     * Replace tagged text content within modified shape
     */
    static replaceText: (replaceText: ReplaceText | ReplaceText[], options?: ReplaceTextOptions) => (element: XmlElement) => void;
    /**
     * Set position and size of modified shape.
     */
    static setPosition: (pos: ShapeCoordinates) => (element: XmlElement) => void;
    /**
     * Update position and size of a shape by a given Value.
     */
    static updatePosition: (pos: ShapeCoordinates) => (element: XmlElement) => void;
    /**
     * Rotate a shape by a given value. Use e.g. 180 to flip a shape.
     * A negative value will rotate counter clockwise.
     * @param degrees Rotate by °
     */
    static rotate: (degrees: number) => (element: XmlElement) => void;
}
