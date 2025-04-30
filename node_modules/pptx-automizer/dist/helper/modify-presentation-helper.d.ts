import IArchive from '../interfaces/iarchive';
import { XmlDocument } from '../types/xml-types';
export default class ModifyPresentationHelper {
    /**
     * Get Collection of slides
     */
    static getSlidesCollection: (xml: XmlDocument) => HTMLCollectionOf<Element>;
    static getSlideMastersCollection: (xml: XmlDocument) => HTMLCollectionOf<Element>;
    /**
     * Pass an array of slide numbers to define a target sort order.
     * First slide starts by 1.
     * @order Array of slide numbers, starting by 1
     */
    static sortSlides: (order: number[]) => (xml: XmlDocument) => void;
    /**
     * Set ids to prevent corrupted pptx.
     * Must start with 256 and increment by one.
     */
    static normalizeSlideIds: (xml: XmlDocument) => void;
    /**
     * Update slideMaster ids to prevent corrupted pptx.
     * - Take first slideMaster id from presentation.xml to start,
     * - then update incremental ids of each p:sldLayoutId in slideMaster[i].xml
     *   (starting by slideMasterId + 1)
     * - and update next slideMaster id with previous p:sldLayoutId + 1
     *
     * p:sldMasterId-ids and p:sldLayoutId-ids need to be in a row, otherwise
     * PowerPoint will complain on any p:sldLayoutId-id lower than its
     * corresponding slideMaster-id. omg.
     */
    static normalizeSlideMasterIds: (xml: XmlDocument, i: number, archive: IArchive) => Promise<void>;
    /**
     * Tracker.files includes all files that have been
     * copied to the root template by automizer. We remove all other files.
     */
    static removeUnusedFiles(xml: XmlDocument, i: number, archive: IArchive): Promise<void>;
    /**
     * PPT won't complain about unused items in [Content_Types].xml,
     * but we remove them anyway in case the file mentioned in PartName-
     * attribute does not exist.
     */
    static removeUnusedContentTypes(xml: XmlDocument, i: number, archive: IArchive): Promise<void>;
    static removedUnusedImages(xml: XmlDocument, i: number, archive: IArchive): Promise<void>;
}
