/**
 * OpenXML constants for DOCX processing
 */

// Twips conversion (1 inch = 1440 twips, 1 cm = 567 twips, 1 pt = 20 twips)
export const TWIPS_PER_PT = 20;
export const TWIPS_PER_CM = 567;

// Font size in half-points (14pt = 28 half-points)
export const FONT_SIZE_14PT = 28;

// Line spacing: 1.1 multiple (base 240 twips per line)
export const LINE_SPACING_1_1 = 264; // 240 * 1.1

// Paragraph spacing
export const SPACE_BEFORE_6PT = 120; // 6 * 20 twips
export const SPACE_AFTER_0PT = 0;

// Indentation
export const RIGHT_INDENT_0_25CM = 142; // ~0.25cm in twips
export const FIRST_LINE_INDENT_1CM = 567; // 1cm = 567 twips

// Default font
export const DEFAULT_FONT = 'Times New Roman';

// OpenXML namespaces
export const NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
};

// Characters to replace
export const EN_DASH = '\u2013';
export const HYPHEN = '-';

// XML element names
export const ELEMENTS = {
  paragraph: 'w:p',
  run: 'w:r',
  text: 'w:t',
  break: 'w:br',
  paragraphProps: 'w:pPr',
  runProps: 'w:rPr',
  spacing: 'w:spacing',
  indent: 'w:ind',
  fonts: 'w:rFonts',
  fontSize: 'w:sz',
  fontSizeCs: 'w:szCs',
  widowControl: 'w:widowControl',
  suppressAutoHyphens: 'w:suppressAutoHyphens',
  contextualSpacing: 'w:contextualSpacing'
};
