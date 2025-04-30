import { XmlElement } from '../types/xml-types';
import { ImageStyle } from '../types/modify-types';
export default class ModifyImageHelper {
    /**
     * Update the "Target" attribute of a created image relation.
     * This will change the image itself. Load images with Automizer.loadMedia
     * @param filename name of target image in root template media folder.
     */
    static setRelationTarget: (filename: string) => (element: XmlElement, arg1: XmlElement) => void;
    static setDuotoneFill: (duotoneParams: ImageStyle['duotone']) => (element: XmlElement) => void;
}
