"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const xml_helper_1 = require("./xml-helper");
const xmldom_1 = require("@xmldom/xmldom");
const dLbl_1 = require("./xml/dLbl");
const lnLRTB_1 = require("./xml/lnLRTB");
class XmlElements {
    constructor(element, params) {
        this.element = element;
        this.document = element.ownerDocument;
        this.params = params;
        this.defaultValues = {
            color: 'CCCCCC',
            size: '1000',
        };
    }
    text() {
        const r = this.document.createElement('a:r');
        r.appendChild(this.textRangeProps());
        r.appendChild(this.textContent());
        let paragraphProps = this.element.getElementsByTagName('a:pPr').item(0);
        if (!paragraphProps) {
            paragraphProps = this.paragraphProps();
        }
        xml_helper_1.XmlHelper.insertAfter(r, paragraphProps);
        return this;
    }
    createTextBody() {
        let txBody = this.element.getElementsByTagName('p:txBody')[0];
        if (!txBody) {
            txBody = this.document.createElement('p:txBody');
            this.element.appendChild(txBody);
            const bodyPr = this.document.createElement('a:bodyPr');
            txBody.appendChild(bodyPr);
            const lstStyle = this.document.createElement('a:lstStyle');
            txBody.appendChild(lstStyle);
            this.paragraphTemplate = this.document.createElement('a:p');
            txBody.appendChild(this.paragraphTemplate);
            this.runTemplate = this.document.createElement('a:r');
            const rPr = this.document.createElement('a:rPr');
            this.runTemplate.appendChild(rPr);
        }
        else {
            let bodyPr = txBody.getElementsByTagName('a:bodyPr')[0];
            if (!bodyPr) {
                bodyPr = this.document.createElement('a:bodyPr');
                txBody.insertBefore(bodyPr, txBody.firstChild);
            }
            let lstStyle = txBody.getElementsByTagName('a:lstStyle')[0];
            if (!lstStyle) {
                lstStyle = this.document.createElement('a:lstStyle');
                txBody.insertBefore(lstStyle, bodyPr.nextSibling);
            }
            const paragraphs = txBody.getElementsByTagName('a:p');
            this.paragraphTemplate = paragraphs[0];
            xml_helper_1.XmlHelper.sliceCollection(paragraphs, 0);
            const runs = this.paragraphTemplate.getElementsByTagName('a:r');
            if (runs.length > 0) {
                this.runTemplate = runs[0];
            }
            else {
                this.runTemplate = this.document.createElement('a:r');
                const rPr = this.document.createElement('a:rPr');
                this.runTemplate.appendChild(rPr);
            }
        }
        return txBody;
    }
    createBodyProperties(txBody) {
        const bodyPr = this.document.createElement('a:bodyPr');
        txBody.appendChild(bodyPr);
        return bodyPr;
    }
    addBulletList(list) {
        const txBody = this.createTextBody();
        this.createBodyProperties(txBody);
        this.processList(txBody, list, 0);
    }
    processList(txBody, items, level) {
        items.forEach((item) => {
            if (Array.isArray(item)) {
                this.processList(txBody, item, level + 1);
            }
            else {
                const p = this.createParagraph(level);
                const r = this.createTextRun(String(item));
                p.appendChild(r);
                txBody.appendChild(p);
            }
        });
    }
    createParagraph(level) {
        const p = this.paragraphTemplate.cloneNode(true);
        const pPr = p.getElementsByTagName('a:pPr')[0];
        if (pPr) {
            if (level > 0) {
                pPr.setAttribute('lvl', String(level));
                pPr.removeAttribute('indent');
                pPr.removeAttribute('marL');
            }
            else {
                pPr.removeAttribute('lvl');
            }
        }
        else {
            const newPPr = this.document.createElement('a:pPr');
            if (level > 0) {
                newPPr.setAttribute('lvl', String(level));
            }
            p.insertBefore(newPPr, p.firstChild);
        }
        const runs = p.getElementsByTagName('a:r');
        xml_helper_1.XmlHelper.sliceCollection(runs, 0);
        return p;
    }
    createTextRun(text) {
        const r = this.runTemplate.cloneNode(true);
        const t = r.getElementsByTagName('a:t')[0];
        if (t) {
            t.textContent = text;
        }
        else {
            const newT = this.document.createElement('a:t');
            newT.textContent = text;
            r.appendChild(newT);
        }
        return r;
    }
    paragraphProps() {
        const p = this.element.getElementsByTagName('a:p').item(0);
        p.appendChild(this.document.createElement('a:pPr'));
        const paragraphRangeProps = this.element
            .getElementsByTagName('a:pPr')
            .item(0);
        const endParaRPr = this.element
            .getElementsByTagName('a:endParaRPr')
            .item(0);
        xml_helper_1.XmlHelper.moveChild(endParaRPr);
        return paragraphRangeProps;
    }
    textRangeProps() {
        const rPr = this.document.createElement('a:rPr');
        const endParaRPr = this.element.getElementsByTagName('a:endParaRPr')[0];
        rPr.setAttribute('lang', endParaRPr.getAttribute('lang'));
        rPr.setAttribute('sz', endParaRPr.getAttribute('sz') || this.defaultValues.size);
        rPr.appendChild(this.line());
        rPr.appendChild(this.effectLst());
        rPr.appendChild(this.lineTexture());
        rPr.appendChild(this.fillTexture());
        return rPr;
    }
    textContent() {
        const t = this.document.createElement('a:t');
        t.textContent = ' ';
        return t;
    }
    effectLst() {
        return this.document.createElement('a:effectLst');
    }
    lineTexture() {
        return this.document.createElement('a:uLnTx');
    }
    fillTexture() {
        return this.document.createElement('a:uFillTx');
    }
    line() {
        const ln = this.document.createElement('a:ln');
        const noFill = this.document.createElement('a:noFill');
        ln.appendChild(noFill);
        return ln;
    }
    solidFill() {
        const solidFill = this.document.createElement('a:solidFill');
        const colorType = this.colorType();
        solidFill.appendChild(colorType);
        return solidFill;
    }
    colorType() {
        var _a, _b;
        const tag = 'a:' + (((_b = (_a = this.params) === null || _a === void 0 ? void 0 : _a.color) === null || _b === void 0 ? void 0 : _b.type) || 'srgbClr');
        const colorType = this.document.createElement(tag);
        this.colorValue(colorType);
        return colorType;
    }
    colorValue(colorType) {
        var _a, _b;
        colorType.setAttribute('val', ((_b = (_a = this.params) === null || _a === void 0 ? void 0 : _a.color) === null || _b === void 0 ? void 0 : _b.value) || this.defaultValues.color);
    }
    dataPoint() {
        const dPt = this.document.createElement('c:dPt');
        dPt.appendChild(this.idx());
        dPt.appendChild(this.spPr());
        const nextSibling = this.element.getElementsByTagName('c:cat')[0];
        if (nextSibling) {
            nextSibling.parentNode.insertBefore(dPt, nextSibling);
        }
        return this;
    }
    spPr() {
        const spPr = this.document.createElement('c:spPr');
        spPr.appendChild(this.solidFill());
        spPr.appendChild(this.line());
        spPr.appendChild(this.effectLst());
        return spPr;
    }
    idx() {
        const idx = this.document.createElement('c:idx');
        idx.setAttribute('val', String(0));
        return idx;
    }
    cellBorder(tag) {
        const border = this.document.createElement(tag);
        border.appendChild(this.solidFill());
        border.appendChild(this.prstDash());
        border.appendChild(this.round());
        border.appendChild(this.lineEnd('headEnd'));
        border.appendChild(this.lineEnd('tailEnd'));
        return this;
    }
    prstDash() {
        const prstDash = this.document.createElement('a:prstDash');
        prstDash.setAttribute('val', 'solid');
        return prstDash;
    }
    round() {
        const round = this.document.createElement('a:round');
        return round;
    }
    lineEnd(type) {
        const lineEnd = this.document.createElement(type);
        lineEnd.setAttribute('type', 'none');
        lineEnd.setAttribute('w', 'med');
        lineEnd.setAttribute('len', 'med');
        return lineEnd;
    }
    shapeProperties() {
        const spPr = this.spPr();
        this.element.appendChild(spPr);
    }
    dataPointLabel() {
        const doc = new xmldom_1.DOMParser().parseFromString(dLbl_1.dLblXml, 'application/xml');
        const ele = doc.getElementsByTagName('c:dLbl')[0];
        const firstChild = this.element.firstChild;
        this.element.insertBefore(ele.cloneNode(true), firstChild);
    }
    tableCellBorder(tag) {
        const doc = new xmldom_1.DOMParser().parseFromString(lnLRTB_1.lnLRTB, 'application/xml');
        const ele = doc.getElementsByTagName(tag)[0];
        const firstChild = this.element.firstChild;
        this.element.insertBefore(ele.cloneNode(true), firstChild);
    }
}
exports.default = XmlElements;
//# sourceMappingURL=xml-elements.js.map