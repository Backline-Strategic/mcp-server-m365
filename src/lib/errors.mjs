/**
 * Actionable error builders.
 * All tool functions catch errors and return these — never throw to MCP caller.
 */

/**
 * Build a standard error response.
 * @param {string} message - What went wrong
 * @param {string} [hint] - How to fix it
 * @param {string} [code] - Error code for programmatic handling
 */
export function makeError(message, hint, code) {
  return {
    success: false,
    error: message,
    ...(hint ? { hint } : {}),
    ...(code ? { code } : {}),
  };
}

/**
 * Wrap a Graph API error with context.
 */
export function graphError(err, context) {
  const msg = err.message ?? String(err);

  // Detect common Graph errors and add hints
  if (msg.includes('403') || msg.includes('Forbidden')) {
    return makeError(
      `${context}: permission denied`,
      'Check that the Azure AD app has the required Graph API permissions and admin consent was granted.',
      'PERMISSION_DENIED'
    );
  }
  if (msg.includes('404') || msg.includes('Not Found')) {
    return makeError(
      `${context}: not found`,
      'The resource may have been deleted or the ID is incorrect.',
      'NOT_FOUND'
    );
  }
  if (msg.includes('401') || msg.includes('Unauthorized')) {
    return makeError(
      `${context}: authentication failed`,
      'Check tenantId, clientId, and clientSecret in your accounts config.',
      'AUTH_FAILED'
    );
  }
  if (msg.includes('throttl') || msg.includes('429') || msg.includes('TooMany')) {
    return makeError(
      `${context}: rate limited`,
      'Graph API throttled this request. Wait a moment and retry.',
      'THROTTLED'
    );
  }

  return makeError(`${context}: ${msg}`);
}

/**
 * Wrap a function call, converting any thrown error to a makeError response.
 */
export async function withErrorHandler(context, fn) {
  try {
    return await fn();
  } catch (err) {
    return graphError(err, context);
  }
}
