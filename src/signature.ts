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
// Single source of truth for the outgoing font, shared by the signature, the
// comment wrapper, and every authored block. Arial is web-safe (present on
// effectively every client, including Outlook desktop's Word engine, which
// ignores web fonts), so recipients see the intended face rather than a
// fallback. Keep this in sync with the OWA "Compose and reply" default font so
// editing a draft in the web UI doesn't visibly reflow it.
const FONT_STACK = 'Arial,Helvetica,sans-serif';
const FONT_SIZE = '11pt';

// Inline style carried by every authored block AND by the wrapper around
// authored content. Deliberately sets NO text color: a fixed color (e.g. #333)
// reads fine on white but stays dark in a client's dark mode, where it can
// vanish against a dark background. Leaving color unset lets the text inherit
// the client's default foreground (near-black in light mode, near-white in
// dark mode).
//
// Deliberately NO max-width, NO opacity, NO margins and NO line-height: earlier
// versions narrowed and dimmed authored comments, and spaced paragraphs with
// margins or a pinned leading, all of which rendered visibly different from
// paragraphs the user types in the OWA composer — a provenance tell.
const BLOCK_STYLE = `font-family:${FONT_STACK}; font-size:${FONT_SIZE}`;
const BODY_STYLE = BLOCK_STYLE;

/**
 * One authored line, in exactly the form the OWA composer writes for a typed
 * line: a <div> carrying the font on the block itself. Matching OWA's own
 * element and style is what keeps authored and hand-typed paragraphs
 * indistinguishable, and it is the reason this module emits <div>s rather
 * than <p>s. When a draft holding <p> blocks is edited in OWA, the editor moves
 * the font-size off the block onto inner <span>s, and the block's line box is
 * then resolved from the reply skeleton's default size — taller leading than
 * the neighbouring typed <div>s. OWA leaves its own <div> structure alone.
 *
 * An empty line holds a lone <br>, which is how OWA represents a blank line.
 */
export function block(inner: string): string {
  return `<div style="${BLOCK_STYLE}">${inner || '<br>'}</div>`;
}

/** A blank line between paragraphs, OWA-style: an empty block holding one <br>. */
export const BLANK_LINE = block('');

// The signature opens with a blank line, exactly as the OWA-configured
// signature block does, so it never butts onto the last line of the message
// regardless of how that message was authored. Content preceding the
// signature must therefore NOT end in its own blank line (see
// normalizeParagraphs and textToBlocks), or the gap doubles.
export const SIGNATURE_HTML =
  '<div id="Signature">' +
  BLANK_LINE +
  `<div style="${BLOCK_STYLE}">` +
  '<span style="opacity:0.6">Geoffrey Challen&nbsp;// </span>' +
  '<a href="https://geoffreychallen.com" title="https://geoffreychallen.com" style="color:rgb(12,100,192)">https://geoffreychallen.com</a>' +
  '</div></div>';

export const SIGNATURE_TEXT = 'Geoffrey Challen // https://geoffreychallen.com';

// Marker used to detect an already-signed HTML body and avoid double-signing.
// Matches the id on SIGNATURE_HTML's wrapper — NOT the bare domain, which also
// appears when the URL is merely mentioned in prose. (An earlier version keyed
// on 'geoffreychallen.com', which meant a comment that spelled out the plain
// signature suppressed the styled one and shipped as unlinked text.)
const SIGNED_MARKER = 'id="Signature"';
const SIGNATURE_OPEN = '<div id="Signature">';

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

/**
 * Render plain text as inline HTML: escape, hyperlink bare URLs, keep line
 * breaks as <br>. Used for quoted history lifted from a text-only reply
 * skeleton, where a single run of lines is the faithful rendering. Authored
 * content goes through textToBlocks instead.
 */
export function textToHtml(text: string): string {
  return linkifyEscapedText(escapeHtml(text)).replace(/\n/g, '<br>');
}

/**
 * Render authored plain text the way OWA renders typed text: one block per
 * line, with blank lines as empty blocks. Trailing blank lines are dropped so
 * the signature's own leading blank line supplies the single gap above it.
 */
export function textToBlocks(text: string): string {
  const trimmed = text.replace(/\r/g, '').replace(/\s+$/, '');
  if (!trimmed) return '';
  return trimmed
    .split('\n')
    .map((line) => block(linkifyEscapedText(escapeHtml(line))))
    .join('');
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
 * Rewrite authored <p> paragraphs into OWA-native blocks: each <p> becomes a
 * font-bearing <div> (see `block`), and the gap between paragraphs becomes a
 * blank-line block rather than a margin. Lists get the same blank line after
 * them that a paragraph does, unless they are nested inside another list item.
 *
 * Why not keep <p> with margins? Graph's reply skeleton carries
 * `<style>p{margin-top:0;margin-bottom:0}</style>`, so bare <p>s collapse into
 * one run-on block; inline margins fix that, but a <p> is still not what the
 * OWA composer writes, and OWA restyles foreign <p>s on edit (font-size moved
 * into spans, block left at the default size), which is the uneven leading the
 * user sees between authored and typed paragraphs. Emitting OWA's own
 * structure sidesteps all of it.
 *
 * Any margin in an author-specified <p> style is dropped (spacing is now the
 * blank line's job); other declarations are kept, with the font added when the
 * author did not set one. The trailing blank line is stripped, and one
 * immediately before the signature is removed, since the signature carries its
 * own.
 */
export function normalizeParagraphs(html: string): string {
  let out = html.replace(/<p(\s[^>]*)?>/gi, (match, attrs?: string) => {
    const a = attrs ?? '';
    const styleMatch = a.match(/\bstyle\s*=\s*("|')([\s\S]*?)\1/i);
    if (!styleMatch) return `<div${a} style="${BLOCK_STYLE}">`;
    const decls = styleMatch[2]
      .split(';')
      .map((d) => d.trim())
      .filter((d) => d && !/^margin/i.test(d));
    if (!decls.some((d) => /^font-family/i.test(d))) decls.push(`font-family:${FONT_STACK}`);
    if (!decls.some((d) => /^font-size/i.test(d))) decls.push(`font-size:${FONT_SIZE}`);
    return `<div${a.replace(styleMatch[0], `style="${decls.join('; ')}"`)}>`;
  });
  out = out.replace(/<\/p>/gi, `</div>${BLANK_LINE}`);
  out = out.replace(/<\/(ul|ol)>(?!\s*<\/(li|ul|ol)>)/gi, (m) => m + BLANK_LINE);
  out = out.split(BLANK_LINE + SIGNATURE_OPEN).join(SIGNATURE_OPEN);
  while (out.endsWith(BLANK_LINE)) out = out.slice(0, -BLANK_LINE.length);
  return out;
}

/**
 * Return a copy of `body` as signed HTML. Every outgoing message goes out as
 * HTML: a plain-text body is converted (escaped, bare URLs hyperlinked, one
 * OWA-style block per line) rather than sent as text, whose URLs render
 * unlinked in most clients. An HTML body has its <p> paragraphs normalized to
 * the same block structure. Any spelled-out plain signature is stripped and
 * replaced by the styled SIGNATURE_HTML. A body already carrying
 * SIGNATURE_HTML is returned unchanged.
 *
 * HTML bodies are also wrapped in an explicit font (FONT_STACK/FONT_SIZE).
 * Without it list items and other non-block markup inherit OWA's base UI font
 * (Segoe UI) when read or edited there — even though the user's compose
 * default is something else, because that default only styles newly typed
 * text, not stored HTML. Stamping the font here makes every outgoing body
 * render in the intended face and keeps OWA from reflowing it on edit.
 */
export function withSignature(body?: MessageBody): MessageBody | undefined {
  if (!body) return body;

  const content = body.content ?? '';
  const contentType = (body.contentType ?? 'text').toLowerCase();

  if (content.includes(SIGNED_MARKER)) {
    return body;
  }

  const rendered =
    contentType === 'html'
      ? normalizeParagraphs(linkifyHtml(content))
      : textToBlocks(stripPlainSignature(content).text);

  if (!rendered) {
    return { contentType: 'html', content: SIGNATURE_HTML };
  }

  const wrapped = `<div style="${BODY_STYLE}">${rendered}</div>`;
  return { contentType: 'html', content: wrapped + SIGNATURE_HTML };
}

/**
 * Stamp the outgoing font onto an HTML body and normalize its paragraphs,
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
    const wrapped = `<div style="${BODY_STYLE}">${textToBlocks(text)}</div>`;
    return {
      contentType: 'html',
      content: stripped ? wrapped + SIGNATURE_HTML : wrapped,
    };
  }

  const content = normalizeParagraphs(linkifyHtml(body.content ?? ''));

  // The author set their own font. Respect it.
  if (/font-family/i.test(content)) return { contentType: 'html', content };

  const open = `<div style="${BODY_STYLE}">`;

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
 * Build the top portion of a reply/forward body: an optional comment followed
 * by the signature, in the user's style. `comment` is treated as plain text
 * unless `opts.raw` is set, in which case it is inserted as-is (already HTML),
 * with its paragraphs normalized to OWA-native blocks.
 */
export function buildReplyTopHtml(comment?: string, opts?: { raw?: boolean }): string {
  if (opts?.raw) {
    const html = comment ?? '';
    // Double-signing guard: a raw HTML comment that already carries the styled
    // signature must not get a second one appended beneath it.
    const signed = html.includes(SIGNED_MARKER);
    const rendered = normalizeParagraphs(linkifyHtml(html));
    const commentHtml = rendered ? `<div style="${BODY_STYLE}">${rendered}</div>` : '';
    return signed ? commentHtml : commentHtml + SIGNATURE_HTML;
  }

  // Plain-text comment: any spelled-out signature is stripped so the styled
  // one below is the only copy, then each line becomes an OWA-style block.
  const rendered = textToBlocks(stripPlainSignature(comment ?? '').text);
  const commentHtml = rendered ? `<div style="${BODY_STYLE}">${rendered}</div>` : '';
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
