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
// near-white in dark mode).
//
// Deliberately NO max-width and NO opacity: earlier versions capped the line
// length (max-width:40em) and dimmed the comment (opacity:0.8), which made
// authored content render visibly different from paragraphs the user types in
// the OWA composer — a provenance tell. Matching OWA-native rendering wins over
// those readability tweaks.
const BODY_STYLE = `font-family:${FONT_STACK}; font-size:${FONT_SIZE}`;

// Block-level style applied to each authored <p>, matching what the OWA
// composer writes for typed paragraphs (`margin-top:1em; margin-bottom:1em`) so
// authored and hand-typed paragraphs are indistinguishable. Symmetric margins
// collapse between adjacent paragraphs, so spacing equals the old bottom-only
// margin while boundary gaps against OWA-typed paragraphs now match too.
//
// font-size is pinned on the block because a paragraph's line box is resolved
// from the *block's* font-size, not an inner <span>'s — when Outlook reflows an
// edited draft it moves font-size onto spans, and an unstyled block would fall
// back to the reply skeleton's default and render taller, uneven leading.
// Deliberately NO explicit line-height: OWA-typed paragraphs use the default
// leading, and pinning a different value (an earlier version used 1.4) made
// authored paragraphs render taller than typed ones.
const PARAGRAPH_BLOCK_STYLE = `margin-top:1em; margin-bottom:1em; font-size:${FONT_SIZE}`;

// Marker used to detect an already-signed HTML body and avoid double-signing.
// Matches the id on SIGNATURE_HTML's wrapper — NOT the bare domain, which also
// appears when the URL is merely mentioned in prose. (An earlier version keyed
// on 'geoffreychallen.com', which meant a comment that spelled out the plain
// signature suppressed the styled one and shipped as unlinked text.)
const SIGNED_MARKER = 'id="Signature"';

interface MessageBody {
  contentType?: string;
  content?: string;
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Matches a bare URL in (possibly entity-escaped) text. Quotes and angle
// brackets terminate the match so hrefs and tags never bleed in.
const URL_PATTERN = /\bhttps?:\/\/[^\s<>"']+/gi;

// Link style shared with SIGNATURE_HTML: an explicit blue affordance that
// clients recognize and adjust for dark mode.
const LINK_STYLE = 'color:rgb(12,100,192)';

/**
 * Wrap bare URLs in styled <a> tags. Input is text that is already
 * HTML-escaped (or a text node lifted out of HTML), so the match can be used
 * directly as both href and label — an escaped `&amp;` is the correct encoding
 * in either position. Trailing punctuation reads as prose, not URL, and is left
 * outside the link.
 */
export function linkifyEscapedText(escaped: string): string {
  return escaped.replace(URL_PATTERN, (url) => {
    const trail = url.match(/[.,;:!?)\]]+$/)?.[0] ?? '';
    const href = trail ? url.slice(0, -trail.length) : url;
    return `<a href="${href}" style="${LINK_STYLE}">${href}</a>${trail}`;
  });
}

/**
 * Auto-link bare URLs in an HTML fragment without touching URLs that are
 * already inside <a> elements, tag attributes, or <style>/<script> content.
 */
export function linkifyHtml(html: string): string {
  let inAnchor = 0;
  let inSkip = 0;
  return html.replace(/<[^>]*>|[^<]+/g, (chunk) => {
    if (chunk.startsWith('<')) {
      const t = chunk.toLowerCase();
      if (/^<a[\s>]/.test(t)) inAnchor++;
      else if (t.startsWith('</a')) inAnchor = Math.max(0, inAnchor - 1);
      else if (/^<(style|script)[\s>]/.test(t)) inSkip++;
      else if (/^<\/(style|script)/.test(t)) inSkip = Math.max(0, inSkip - 1);
      return chunk;
    }
    return inAnchor || inSkip ? chunk : linkifyEscapedText(chunk);
  });
}

/** Render plain text as HTML: escape, hyperlink bare URLs, keep line breaks. */
export function textToHtml(text: string): string {
  return linkifyEscapedText(escapeHtml(text)).replace(/\n/g, '<br>');
}

/**
 * Remove any spelled-out plain-text signature from authored text. Callers
 * (humans and models alike) habitually type the signature at the end of a
 * message; the styled SIGNATURE_HTML is appended in its place, so leaving the
 * plain copy in would double-sign — and, worse, ship the URL as unlinked text.
 */
function stripPlainSignature(text: string): { text: string; stripped: boolean } {
  if (!text.includes(SIGNATURE_TEXT)) return { text, stripped: false };
  return {
    text: text.split(SIGNATURE_TEXT).join('').replace(/\s+$/, ''),
    stripped: true,
  };
}

/**
 * Return a copy of `body` as signed HTML. Every outgoing message goes out as
 * HTML: a plain-text body is converted (escaped, bare URLs hyperlinked, line
 * breaks preserved) rather than sent as text, whose URLs render unlinked in
 * most clients. Any spelled-out plain signature is stripped and replaced by
 * the styled SIGNATURE_HTML. A body already carrying SIGNATURE_HTML is
 * returned unchanged.
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

  if (content.includes(SIGNED_MARKER)) {
    return body;
  }

  // A converted text body is <br>-joined lines with no block margins, so add a
  // blank line before the signature; authored HTML's last <p> margin (see the
  // spacer rationale atop this file) already supplies that gap.
  const rendered =
    contentType === 'html'
      ? linkifyHtml(content)
      : textToHtml(stripPlainSignature(content).text).concat('<br>');

  if (!rendered || rendered === '<br>') {
    return { contentType: 'html', content: SIGNATURE_HTML };
  }

  const wrapped = `<div style="font-family:${FONT_STACK}; font-size:${FONT_SIZE}">${rendered}</div>`;
  return { contentType: 'html', content: wrapped + SIGNATURE_HTML };
}

/**
 * Stamp the outgoing font onto an HTML body and normalize paragraph spacing,
 * without touching the signature.
 *
 * `withSignature` does this on the create path, but a PATCH that replaces a
 * draft's body used to go to Graph untouched — so a body authored as bare <p>
 * tags inherited OWA's Segoe UI instead of Arial, and the paragraphs collapsed
 * against the reply skeleton's margin reset. Replacing a body is the normal way
 * to revise a draft, and it should not silently produce different typography
 * from the message that draft started as.
 *
 * Idempotent in the ways that matter. A body that already declares a
 * font-family is left to its author. A full <html>…<body> document gets the
 * wrapper inside <body> rather than around the document, which would be
 * malformed.
 *
 * Like every other outgoing path, the result is always HTML with bare URLs
 * hyperlinked: a plain-text replacement body is converted rather than stored
 * as text. The signature stays the caller's responsibility on this path (a
 * PATCH replaces the whole body, which normally already carries it) — but a
 * spelled-out plain-text signature is upgraded to the styled SIGNATURE_HTML so
 * it keeps its hyperlink.
 */
export function withOutgoingFont(body?: MessageBody): MessageBody | undefined {
  if (!body) return body;

  if ((body.contentType ?? 'text').toLowerCase() !== 'html') {
    const { text, stripped } = stripPlainSignature(body.content ?? '');
    const rendered = textToHtml(text);
    const wrapped = `<div style="font-family:${FONT_STACK}; font-size:${FONT_SIZE}">${rendered}${stripped ? '<br>' : ''}</div>`;
    return {
      contentType: 'html',
      content: stripped ? wrapped + SIGNATURE_HTML : wrapped,
    };
  }

  const content = ensureParagraphSpacing(linkifyHtml(body.content ?? ''));

  // The author set their own font. Respect it.
  if (/font-family/i.test(content)) return { contentType: 'html', content };

  const open = `<div style="font-family:${FONT_STACK}; font-size:${FONT_SIZE}">`;

  const bodyOpen = content.match(/<body[^>]*>/i);
  if (bodyOpen && bodyOpen.index !== undefined) {
    const start = bodyOpen.index + bodyOpen[0].length;
    const closeAt = content.search(/<\/body>/i);
    const end = closeAt > -1 ? closeAt : content.length;
    return {
      contentType: 'html',
      content:
        content.slice(0, start) + open + content.slice(start, end) + '</div>' + content.slice(end),
    };
  }

  return { contentType: 'html', content: open + content + '</div>' };
}

/**
 * Graph's reply/forward skeleton carries `<style>p{margin-top:0;margin-bottom:0}</style>`
 * in its <head>, which collapses the vertical spacing between <p> paragraphs in an
 * injected HTML comment so they render as one run-on block. Inline styles win over that
 * <style> reset in Outlook and other clients, so give every bare <p> the block-level
 * PARAGRAPH_BLOCK_STYLE (OWA-style symmetric margins for spacing, plus font-size so
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
      const merged = `${styleMatch[2].trim().replace(/;\s*$/, '')}; margin-top:1em; margin-bottom:1em`;
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
  if (opts?.raw) {
    const html = comment ?? '';
    // Double-signing guard: a raw HTML comment that already carries the styled
    // signature must not get a second one appended beneath it.
    const signed = html.includes(SIGNED_MARKER);
    // No spacer here: the trailing <p> already carries PARAGRAPH_BLOCK_STYLE's
    // 1em bottom margin, and a <br> would double the gap.
    const rendered = ensureParagraphSpacing(linkifyHtml(html));
    const commentHtml = html ? `<div style="${BODY_STYLE}">${rendered}</div>` : '';
    return signed ? commentHtml : commentHtml + SIGNATURE_HTML;
  }

  // Plain-text comment: any spelled-out signature is stripped so the styled
  // one below is the only copy, then the text is escaped with bare URLs
  // hyperlinked. The <br>-joined lines end flush against the wrapper, so a
  // spacer keeps the signature from butting onto the last line.
  const { text } = stripPlainSignature(comment ?? '');
  const commentHtml = text ? `<div style="${BODY_STYLE}">${textToHtml(text)}<br></div>` : '';
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
  const top = buildReplyTopHtml(comment, opts);

  // A text skeleton (Graph produces one when the original message was plain
  // text) is converted so the reply still goes out as HTML with a styled,
  // hyperlinked signature; the quoted history keeps its line breaks and gets
  // its URLs linked like everything else.
  if (contentType !== 'html') {
    const quoted = content
      ? `<div style="font-family:${FONT_STACK}; font-size:${FONT_SIZE}">${textToHtml(content)}</div>`
      : '';
    return { contentType: 'html', content: top + quoted };
  }
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
