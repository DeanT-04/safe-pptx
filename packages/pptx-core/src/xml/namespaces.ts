/** OOXML namespace URIs used by PresentationML packages. */
export const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  dc: 'http://purl.org/dc/elements/1.1/',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  p15: 'http://schemas.microsoft.com/office/powerpoint/2012/main',
  p18: 'http://schemas.microsoft.com/office/powerpoint/2018/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
} as const;

export type NsPrefix = keyof typeof NS;

export function nsUri(prefix: NsPrefix): string {
  return NS[prefix];
}

/** Clark-form qualified name, e.g. qn('p','sp') → '{…presentationml…}sp'. */
export function qn(prefix: NsPrefix, local: string): string {
  return `{${NS[prefix]}}${local}`;
}
