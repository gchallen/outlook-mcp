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

// Shared inline style matching the user's body/signature text.
const BODY_STYLE = 'font-family:Arial,Helvetica,sans-serif; font-size:11pt; color:rgb(51,51,51)';

// Marker used to detect an already-signed body and avoid double-signing.
const SIGNATURE_MARKER = 'geoffreychallen.com';

interface MessageBody {
  contentType?: string;
  content?: string;
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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

/**
 * Graph's reply/forward skeleton carries `<style>p{margin-top:0;margin-bottom:0}</style>`
 * in its <head>, which collapses the vertical spacing between <p> paragraphs in an
 * injected HTML comment so they render as one run-on block. Inline styles win over that
 * <style> reset in Outlook and other clients, so give every <p> that lacks an explicit
 * margin a bottom margin. This keeps authored `<p>`-per-paragraph HTML rendering with
 * real paragraph breaks without the caller hand-inserting blank `<div><br></div>` lines.
 */
export function ensureParagraphSpacing(html: string): string {
  return html.replace(/<(p)(\s[^>]*)?>/gi, (match, tag: string, attrs?: string) => {
    const a = attrs ?? '';
    const styleMatch = a.match(/\bstyle\s*=\s*("|')([\s\S]*?)\1/i);
    if (styleMatch) {
      // Respect an author-specified margin; otherwise merge one into the existing style.
      if (/margin/i.test(styleMatch[2])) return match;
      const merged = `${styleMatch[2].trim().replace(/;\s*$/, '')}; margin:0 0 1em 0`;
      return `<${tag}${a.replace(styleMatch[0], `style="${merged}"`)}>`;
    }
    return `<${tag}${a} style="margin:0 0 1em 0">`;
  });
}

/**
 * Build the top portion of a reply/forward body: an optional comment followed
 * by the signature, in the user's style. `comment` is treated as plain text
 * unless `opts.raw` is set, in which case it is inserted as-is (already HTML),
 * with paragraph spacing normalized so <p> blocks survive the reply skeleton's
 * margin reset.
 */
export function buildReplyTopHtml(comment?: string, opts?: { raw?: boolean }): string {
  const inner = opts?.raw
    ? ensureParagraphSpacing(comment ?? '')
    : escapeHtml(comment ?? '').replace(/\n/g, '<br>');
  const commentHtml = comment ? `<div style="${BODY_STYLE}">${inner}</div>` : '';
  return commentHtml + SIGNATURE_HTML;
}

/**
 * Inject comment + signature into the draft body produced by Graph's
 * createReply/createReplyAll/createForward (which contains the quoted thread).
 * The new content is placed at the top, above the quoted history.
 */
export function injectReplyTop(
  draftBody: MessageBody | undefined,
  comment?: string,
  opts?: { raw?: boolean },
): MessageBody {
  const content = draftBody?.content ?? '';
  const contentType = (draftBody?.contentType ?? 'html').toLowerCase();

  if (contentType !== 'html') {
    const top = comment ? `${comment}\n\n${SIGNATURE_TEXT}` : SIGNATURE_TEXT;
    return { contentType: draftBody?.contentType ?? 'text', content: `${top}\n\n${content}` };
  }

  const top = buildReplyTopHtml(comment, opts);
  const bodyOpen = content.match(/<body[^>]*>/i);
  if (bodyOpen && bodyOpen.index !== undefined) {
    const insertAt = bodyOpen.index + bodyOpen[0].length;
    return {
      contentType: 'html',
      content: content.slice(0, insertAt) + top + content.slice(insertAt),
    };
  }

  return { contentType: 'html', content: top + content };
}
