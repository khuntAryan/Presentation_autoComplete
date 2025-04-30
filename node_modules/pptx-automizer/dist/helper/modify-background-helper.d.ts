import { XmlDocument } from '../types/xml-types';
import { Color } from '../types/modify-types';
import { IMaster } from '../index';
export default class ModifyBackgroundHelper {
    /**
     * Set solid fill of master background
     */
    static setSolidFill: (color: Color) => (slideMasterXml: XmlDocument) => void;
    /**
     * Modify a slideMaster background image's relation target
     * @param master
     * @param imageName
     */
    static setRelationTarget: (master: IMaster, imageName: string) => void;
    /**
     * Extract background properties from slideMaster xml
     */
    static getBackgroundProperties: (slideMasterXml: XmlDocument) => string;
}
