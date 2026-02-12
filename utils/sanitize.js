/**
 * Sanitizes scope text to remove any pricing, quantities, or rates
 * that should never appear in our scope-only format.
 */

const BLOCKED_PATTERNS = [
  // Currency amounts
  /\$\s*[\d,]+(\.\d{1,2})?/g,
  /USD\s*[\d,]+/gi,
  /[\d,]+\s*dollars?/gi,

  // Unit rates
  /\/\s*(SF|LF|SY|CY|EA|each|unit|ton|lb|ft|inch|in\b)/gi,
  /per\s+(SF|LF|SY|CY|EA|each|unit|ton|lb|ft|lineal\s+foot)/gi,
  /@\s*\$[\d,.]+/g,

  // Totals / lump sum phrases
  /\btotal\s+cost\b/gi,
  /\blump\s+sum\b/gi,
  /\ballowance\b/gi,
  /\bunit\s+price\b/gi,
  /\bbid\s+price\b/gi,
  /\bcontract\s+amount\b/gi,

  // Quantity patterns (number + unit abbreviation)
  /\b\d+\s*(SF|LF|SY|CY|EA|tons?|lbs?|gallons?|sq\s*ft|lin\s*ft)\b/gi,
];

/**
 * Sanitizes a single scope string, removing pricing/quantity fragments.
 * @param {string} text
 * @returns {string}
 */
export function sanitizeScope(text) {
  if (!text || typeof text !== 'string') return text;

  let cleaned = text;
  for (const pattern of BLOCKED_PATTERNS) {
    cleaned = cleaned.replace(pattern, '[verify]');
  }

  // Collapse any double spaces from removals
  cleaned = cleaned.replace(/\s{2,}/g, ' ').trim();

  return cleaned;
}

/**
 * Sanitizes all scope_points in a divisions array.
 * @param {Array} divisions
 * @returns {Array}
 */
export function sanitizeDivisions(divisions) {
  return divisions.map(div => ({
    ...div,
    items: div.items.map(item => ({
      ...item,
      scope_points: item.scope_points.map(sanitizeScope),
    })),
  }));
}
