/**
 * Token-efficient response formatters.
 * Strips null fields, truncates previews, adds pagination metadata.
 */

/**
 * Remove null/undefined fields to save tokens.
 */
export function compact(obj) {
  if (Array.isArray(obj)) return obj.map(compact);
  if (obj && typeof obj === 'object') {
    return Object.fromEntries(
      Object.entries(obj)
        .filter(([, v]) => v != null)
        .map(([k, v]) => [k, compact(v)])
    );
  }
  return obj;
}

/**
 * Truncate a string preview.
 */
export function preview(text, maxLen = 150) {
  if (!text) return undefined;
  const clean = text.replace(/<[^>]+>/g, '').trim();
  return clean.length <= maxLen ? clean : clean.slice(0, maxLen) + '…';
}

/**
 * Wrap items in a paginated response envelope.
 */
export function paginated(items, total, limit) {
  return {
    count: items.length,
    hasMore: total != null ? total > items.length : items.length >= limit,
    items,
  };
}

/**
 * Normalize a Graph dateTime string to ISO (appends Z if missing).
 */
export function toUtcIso(dateTime) {
  if (!dateTime) return undefined;
  return dateTime.endsWith('Z') ? dateTime : dateTime + 'Z';
}
