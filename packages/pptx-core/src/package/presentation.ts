import { PptxZip } from '../zip/pptx_zip.js';
import { ContentTypes } from './content_types.js';
import { RelResolver } from './rels.js';
import { childrenByName, getAttr } from '../xml/parse.js';
import type { Document } from '../xml/parse.js';

export const PRESENTATION_PART = 'ppt/presentation.xml';

export interface SlideRef {
  /** 0-based position in `sldIdLst` — the display order, not the filename order. */
  index: number;
  /** `p:sldId/@id` — must be ≥256 and unique across top-level id lists. */
  sldId: number;
  relId: string;
  partPath: string;
}

export interface PackagePartRefs {
  slideMasters: string[];
  notesMasters: string[];
  slideLayouts: string[];
  notesSlides: string[];
  themeParts: string[];
}

/**
 * The loaded .pptx package: zip + content types + relationship resolution +
 * slide order. All structural reads/writes go through this model.
 */
export class PptxPackage {
  readonly rels: RelResolver;
  readonly contentTypes: ContentTypes;
  slides: SlideRef[] = [];
  readonly partRefs: PackagePartRefs = {
    slideMasters: [],
    notesMasters: [],
    slideLayouts: [],
    notesSlides: [],
    themeParts: [],
  };

  private constructor(readonly zip: PptxZip) {
    this.contentTypes = ContentTypes.parse(zip.doc('[Content_Types].xml'));
    this.rels = new RelResolver((relsPath) =>
      zip.hasPart(relsPath) ? zip.doc(relsPath) : null,
    );
  }

  static async load(buffer: Buffer): Promise<PptxPackage> {
    const zip = await PptxZip.load(buffer);
    const pkg = new PptxPackage(zip);
    pkg.loadSlideList();
    pkg.loadPartRefs();
    return pkg;
  }

  private loadSlideList(): void {
    const presDoc = this.zip.doc(PRESENTATION_PART);
    const sldIdLst = childrenByName(presDoc.documentElement, 'p', 'sldIdLst')[0];
    if (!sldIdLst) {
      throw new Error('presentation.xml has no p:sldIdLst — invalid PresentationML');
    }
    this.slides = childrenByName(sldIdLst, 'p', 'sldId').map((el, index) => {
      const sldId = Number(el.getAttribute('id'));
      const relId = getAttr(el, 'r', 'id') ?? '';
      const partPath = this.rels.targetOf(PRESENTATION_PART, relId);
      if (partPath === null || !this.zip.hasPart(partPath)) {
        throw new Error(
          `slide ${index + 1}: relationship ${relId} does not resolve to an existing part` +
            (partPath ? ` (got ${partPath})` : ''),
        );
      }
      return { index, sldId, relId, partPath };
    });
  }

  private loadPartRefs(): void {
    this.partRefs.slideMasters = this.rels.targetsOfType(PRESENTATION_PART, 'slideMaster');
    this.partRefs.notesMasters = this.rels.targetsOfType(PRESENTATION_PART, 'notesMaster');
    const layouts = new Set<string>();
    const notesSlides = new Set<string>();
    const themes = new Set<string>();
    for (const masterPath of this.partRefs.slideMasters) {
      this.rels.targetsOfType(masterPath, 'slideLayout').forEach((p) => layouts.add(p));
      this.rels.targetsOfType(masterPath, 'theme').forEach((p) => themes.add(p));
    }
    for (const slidePath of this.slides.map((s) => s.partPath)) {
      this.rels.targetsOfType(slidePath, 'slideLayout').forEach((p) => layouts.add(p));
      this.rels.targetsOfType(slidePath, 'notesSlide').forEach((p) => notesSlides.add(p));
    }
    for (const notesMasterPath of this.partRefs.notesMasters) {
      this.rels.targetsOfType(notesMasterPath, 'theme').forEach((p) => themes.add(p));
    }
    this.partRefs.slideLayouts = [...layouts].sort();
    this.partRefs.notesSlides = [...notesSlides].sort();
    this.partRefs.themeParts = [...themes].sort();
  }

  /** Reload structural indexes after slide add/delete/reorder. */
  refreshStructure(): void {
    this.rels.invalidateAll();
    this.loadSlideList();
    this.loadPartRefs();
  }

  slideByIndex(index: number): SlideRef | undefined {
    return this.slides[index];
  }

  slideByPartPath(partPath: string): SlideRef | undefined {
    return this.slides.find((s) => s.partPath === partPath);
  }

  notesSlideForSlide(slidePartPath: string): string | null {
    return this.rels.targetsOfType(slidePartPath, 'notesSlide')[0] ?? null;
  }

  layoutForSlide(slidePartPath: string): string | null {
    return this.rels.targetsOfType(slidePartPath, 'slideLayout')[0] ?? null;
  }

  masterForLayout(layoutPartPath: string): string | null {
    return this.rels.targetsOfType(layoutPartPath, 'slideMaster')[0] ?? null;
  }

  presentationDoc(): Document {
    return this.zip.doc(PRESENTATION_PART);
  }
}
