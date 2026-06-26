// The user's email signature, extracted from their sent mail. Appended
// automatically to outgoing messages. This MCP server is single-user.

export const SIGNATURE_HTML =
  '<div id="Signature">' +
  '<div style="font-family:Arial,Helvetica,sans-serif; font-size:11pt; color:rgb(51,51,51)"><br></div>' +
  '<div style="font-family:Arial,Helvetica,sans-serif; font-size:11pt; color:rgb(51,51,51)">' +
  '<span style="color:rgb(102,102,102)">Geoffrey Challen&nbsp;// </span>' +
  '<span style="color:rgb(12,100,192)">' +
  '<a href="https://geoffreychallen.com" title="https://geoffreychallen.com" style="color:rgb(12,100,192)">https://geoffreychallen.com</a>' +
  '</span></div></div>';

export const SIGNATURE_TEXT = 'Geoffrey Challen // https://geoffreychallen.com';

// Marker used to detect an already-signed body and avoid double-signing.
const SIGNATURE_MARKER = 'geoffreychallen.com';

interface MessageBody {
  contentType?: string;
  content?: string;
}

/**
 * Return a copy of `body` with the signature appended. HTML bodies get the
 * styled HTML signature; text bodies get the plain-text version. A body that
 * already contains the signature is returned unchanged.
 */
export function withSignature(body?: MessageBody): MessageBody | undefined {
  if (!body) return body;

  const content = body.content ?? '';
  const contentType = (body.contentType ?? 'text').toLowerCase();

  if (content.includes(SIGNATURE_MARKER)) {
    return body;
  }

  if (contentType === 'html') {
    return { contentType: 'html', content: content + SIGNATURE_HTML };
  }

  return {
    contentType: body.contentType ?? 'text',
    content: content ? `${content}\n\n${SIGNATURE_TEXT}` : SIGNATURE_TEXT,
  };
}
