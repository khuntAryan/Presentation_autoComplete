import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { XmlElement } from '../types/xml-types';
type Expressions = {
    openingTag: string;
    closingTag: string;
};
type CharacterSplit = {
    from: number;
    to: number;
    text: string;
};
export default class TextReplaceHelper {
    expressions: Expressions;
    element: XmlElement;
    newNodes: XmlElement[];
    options: ReplaceTextOptions;
    constructor(options: ReplaceTextOptions, element: XmlElement);
    isolateTaggedNodes(): this;
    splitTextBlock(block: XmlElement, matches: RegExpMatchArray[], textContent: string): void;
    getCharacterSplit(matches: RegExpMatchArray[], textContent: string): CharacterSplit[];
    pushCharacterSplit(split: CharacterSplit[], from: number, to: number, text: string): void;
    insertBlock(block: XmlElement, text: string): XmlElement;
    applyReplacements(replaceTexts: ReplaceText[]): void;
    applyReplacement(replaceText: ReplaceText, textBlock: XmlElement, currentIndex: number): void;
    assertTextBlocks(length: number, textBlock: any): XmlElement[];
    updateTextNode(textNode: XmlElement, sourceText: any, replace: any, by: any): void;
    getTextElement(block: XmlElement): XmlElement;
    getRegExp(): RegExp;
}
export {};
