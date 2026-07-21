/**
 * JBA OS SHARED UTILITIES
 *
 * Keep reusable, project-wide helper functions in this file.
 * Do not duplicate these function names in other .gs files.
 */

/**
 * Escapes a value before inserting it into an HTML email or HTML fragment.
 */
function escapeHtml_(value) {
  if (value === null || value === undefined) return '';

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
