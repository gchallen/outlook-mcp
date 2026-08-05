// The user's email signature, extracted from their sent mail. Appended
// automatically to outgoing messages. This MCP server is single-user.
//
// Dark-mode-safe styling (same rationale as the comment wrapper below): the
// muted "Geoffrey Challen //" name sets NO color and uses opacity:0.6 to dim
// against whatever background is behind it — 0.6 over white reproduces the old
// rgb(102,102,102) gray, and it stays legible in dark mode instead of going
// dim-on-dark. The URL keeps an explicit blue: a link needs a recognizable
// color affordance, and clients expect/adjust link colors. Outlook desktop
// ignores opacity and shows full-contrast default text — a safe degradation.
//
// No leading blank-line spacer: the body's last paragraph already carries a 1em
// bottom margin (PARAGRAPH_BLOCK_STYLE), so a <br> here would stack a second empty
// line above the signature. The margin supplies the single gap on its own.
// Single source of truth for the outgoing font, shared by the signature, the
// comment wrapper, and each authored <p>. Arial is web-safe (present on
// effectively every client, including Outlook desktop's Word engine, which
// ignores web fonts), so recipients see the intended face rather than a
// fallback. Keep this in sync with the OWA "Compose and reply" default font so
// editing a draft in the web UI doesn't visibly reflow it.
const FONT_STACK = 'Arial,Helvetica,sans-serif';
const FONT_SIZE = '11pt';

export const SIGNATURE_HTML =
  '<div id="Signature">' +
  `<div style="font-family:${FONT_STACK}; font-size:${FONT_SIZE}">` +
  '<span style="opacity:0.6">Geoffrey Challen&nbsp;// </span>' +
  '<a href="https://geoffreychallen.com" title="https://geoffreychallen.com" style="color:rgb(12,100,192)">https://geoffreychallen.com</a>' +
  '</div></div>';

export const SIGNATURE_TEXT = 'Geoffrey Challen // https://geoffreychallen.com';

// Inline style for the comment wrapper. Deliberately sets NO text color: a
// fixed color (e.g. #333) reads fine on white but stays dark in a client's dark
// mode, where it can vanish against a dark background. Leaving color unset lets
// the text inherit the client's default foreground (near-black in light mode,
// near-white in dark mode). `opacity:0.8` then dims it relative to whatever
// background is behind it, so it reads slightly softer in BOTH modes — and 0.8
// over white is identical to the old #333. Outlook desktop ignores opacity and
// simply shows full-contrast default text, which is a safe degradation.
// `max-width` caps the line length for readability on wide screens; clients that
// don't honor it fall back to full width.
const BODY_STYLE = `font-family:${FONT_STACK}; font-size:${FONT_SIZE}; max-width:40em; opacity:0.8`;

// Block-level style applied to each authored <p>. A paragraph's line-height is
// resolved from the *block's* font-size, not an inner <span>'s — and when Outlook
// reflows an edited draft it tends to move font-size onto inline spans, leaving the
// <p> block to fall back to the reply's default and render taller, uneven leading.
// Pinning font-size AND an explicit line-height on the block keeps the leading
// deterministic and consistent across paragraphs regardless of that reflow.
const PARAGRAPH_BLOCK_STYLE = `margin:0 0 1em 0; font-size:${FONT_SIZE}; line-height:1.4`;

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
 *
 * HTML bodies are also wrapped in an explicit font (FONT_STACK/FONT_SIZE).
 * Without it the authored markup carries no font-family and inherits OWA's base
 * UI font (Segoe UI) when read or edited there — even though the user's compose
 * default is something else, because that default only styles newly typed text,
 * not stored HTML. Stamping the font here makes every outgoing body render in
 * the intended face and keeps OWA from reflowing it on edit.
 */
export function withSignature(body?: MessageBody): MessageBody | undefined {
  if (!body) return body;

  const content = body.content ?? '';
  const contentType = (body.contentType ?? 'text').toLowerCase();

  if (content.includes(SIGNATURE_MARKER)) {
    return body;
  }

  if (contentType === 'html') {
    const wrapped = `<div style="font-family:${FONT_STACK}; font-size:${FONT_SIZE}">${content}</div>`;
    return { contentType: 'html', content: wrapped + SIGNATURE_HTML };
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
 * <style> reset in Outlook and other clients, so give every bare <p> the block-level
 * PARAGRAPH_BLOCK_STYLE (bottom margin for spacing, plus font-size and line-height so
 * leading is deterministic and consistent). This keeps authored `<p>`-per-paragraph HTML
 * rendering with real breaks and even leading, no hand-inserted `<div><br></div>` lines.
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
    return `<${tag}${a} style="${PARAGRAPH_BLOCK_STYLE}">`;
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
  // Same double-signing guard as withSignature(): a comment that already carries
  // the signature must not get a second one appended beneath it.
  const signed = (comment ?? '').includes(SIGNATURE_MARKER);

  // A plain-text comment is rendered as <br>-joined lines, so its wrapper ends
  // flush against the last line and the signature would butt directly onto it.
  // Raw HTML comments don't need this: their trailing <p> already carries
  // PARAGRAPH_BLOCK_STYLE's 1em bottom margin, and a <br> would double the gap.
  const needsSpacer = !opts?.raw && !!comment && !signed;

  const rendered = opts?.raw
    ? ensureParagraphSpacing(comment ?? '')
    : escapeHtml(comment ?? '').replace(/\n/g, '<br>');
  const inner = needsSpacer ? `${rendered}<br>` : rendered;

  const commentHtml = comment ? `<div style="${BODY_STYLE}">${inner}</div>` : '';
  if (signed) return commentHtml;
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
    const top =
      comment && comment.includes(SIGNATURE_MARKER)
        ? comment
        : comment
          ? `${comment}\n\n${SIGNATURE_TEXT}`
          : SIGNATURE_TEXT;
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
