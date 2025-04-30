import { ModifyTableParams, TableData } from '../types/table-types';
import { XmlDocument, XmlElement } from '../types/xml-types';
export default class ModifyTableHelper {
    static setTable: (data: TableData, params?: ModifyTableParams) => (element: XmlElement) => void;
    static setTableData: (data: TableData) => (element: XmlDocument | XmlElement) => void;
    static adjustHeight: (data: TableData) => (element: XmlDocument | XmlElement) => void;
    static adjustWidth: (data: TableData) => (element: XmlDocument | XmlElement) => void;
    static updateColumnWidth: (index: number, size: number) => (element: XmlDocument | XmlElement) => void;
    static updateRowHeight: (index: number, size: number) => (element: XmlDocument | XmlElement) => void;
}
