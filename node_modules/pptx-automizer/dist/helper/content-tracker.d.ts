import { Target, TrackedFiles, TrackedRelation, TrackedRelationInfo, TrackedRelations, TrackedRelationTag } from '../types/types';
import { RelationshipAttribute } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
export declare class ContentTracker {
    archive: IArchive;
    files: TrackedFiles;
    relations: TrackedRelations;
    relationTags: TrackedRelationTag[];
    constructor();
    reset(): void;
    trackFile(file: string): void;
    trackRelation(file: string, attributes: RelationshipAttribute): void;
    analyzeContents(archive: IArchive): Promise<void>;
    setArchive(archive: IArchive): void;
    /**
     * This will be replaced by future slideMaster handling.
     */
    trackSlideMasters(): Promise<void>;
    trackSlideLayouts(): Promise<void>;
    addAndAnalyze(trackedRelations: TrackedRelation[], section: string): Promise<void>;
    getRelatedContents(trackedRelations: TrackedRelation[]): Promise<Target[]>;
    getRelationTag(source: string): TrackedRelationTag;
    analyzeRelationships(): Promise<void>;
    analyzeRelationship(relationTagInfo: TrackedRelationTag): Promise<void>;
    pushRelationTagTargets(file: string, filename: string, relationTag: TrackedRelation, relationTagInfo: any): Promise<void>;
    addCreatedRelationsFunctions(addTargets: Target[], createdRelations: TrackedRelationInfo[], relationTagInfo: TrackedRelationTag): void;
    getCreatedContent(createdRelations: TrackedRelationInfo[], addTarget: Target): () => TrackedRelationInfo;
    addRelatedContent(relationTagInfo: TrackedRelationTag, addTarget: Target): () => Promise<Target>;
    collect(section: string, role: string, collection?: string[]): Promise<string[]>;
    filterRelations(section: string, target: string): TrackedRelationInfo[];
}
export declare const contentTracker: ContentTracker;
