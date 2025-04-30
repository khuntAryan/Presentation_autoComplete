"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const modify_table_1 = require("../modify/modify-table");
const xml_slide_helper_1 = require("./xml-slide-helper");
class ModifyTableHelper {
}
exports.default = ModifyTableHelper;
ModifyTableHelper.setTable = (data, params) => (element) => {
    const modTable = new modify_table_1.ModifyTable(element, data);
    if (params === null || params === void 0 ? void 0 : params.expand) {
        params === null || params === void 0 ? void 0 : params.expand.forEach((expand) => {
            const tableInfo = xml_slide_helper_1.XmlSlideHelper.readTableInfo(element);
            const targetCell = tableInfo.find((infoCell) => infoCell.textContent === expand.tag);
            if (targetCell) {
                if (expand.mode === 'row') {
                    modTable.expandRows(expand.count, targetCell.row);
                }
                else {
                    if (targetCell.gridSpan) {
                        modTable.expandSpanColumns(expand.count, targetCell.column, targetCell.gridSpan);
                    }
                    else {
                        modTable.expandColumns(expand.count, targetCell.column);
                    }
                }
            }
        });
    }
    modTable.modify(params);
    if (params === null || params === void 0 ? void 0 : params.setHeight) {
        modTable.setSize('cy', params.setHeight);
    }
    if (params === null || params === void 0 ? void 0 : params.setWidth) {
        modTable.setSize('cx', params.setWidth);
    }
    if (!params || (params === null || params === void 0 ? void 0 : params.adjustHeight)) {
        modTable.adjustHeight();
    }
    if (!params || (params === null || params === void 0 ? void 0 : params.adjustWidth)) {
        modTable.adjustWidth();
    }
};
ModifyTableHelper.setTableData = (data) => (element) => {
    const modTable = new modify_table_1.ModifyTable(element, data);
    modTable.modify();
};
ModifyTableHelper.adjustHeight = (data) => (element) => {
    const modTable = new modify_table_1.ModifyTable(element, data);
    modTable.adjustHeight();
};
ModifyTableHelper.adjustWidth = (data) => (element) => {
    const modTable = new modify_table_1.ModifyTable(element, data);
    modTable.adjustWidth();
};
ModifyTableHelper.updateColumnWidth = (index, size) => (element) => {
    const modTable = new modify_table_1.ModifyTable(element);
    modTable.updateColumnWidth(index, size);
};
ModifyTableHelper.updateRowHeight = (index, size) => (element) => {
    const modTable = new modify_table_1.ModifyTable(element);
    modTable.updateRowHeight(index, size);
};
//# sourceMappingURL=modify-table-helper.js.map