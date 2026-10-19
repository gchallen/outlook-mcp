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

// A plain-text list item: "- ", "* " or "• " for bullets, "1. " or "1) " for
// numbers. The space after the marker is required, so "-5 degrees" and
// "3.14" stay prose, and numbers stop at two digits so a line opening with a
// year ("2026. A year of ...") does too.
const BULLET_ITEM = /^\s*[-*•]\s+(.*)$/;
const NUMBERED_ITEM = /^\s*(\d{1,2})[.)]\s+(.*)$/;

/**
 * Render authored plain text the way OWA renders typed text: one block per
 * line, with blank lines as empty blocks. Consecutive bullet lines become a
 * <ul> and consecutive numbered lines an <ol> (starting at the first number),
 * with the font stamped on each list element as normalizeParagraphs does for
 * authored HTML. Callers write lists in plain text far more often than they
 * remember to switch to HTML, and a literal "- " reads as a stray hyphen.
 * Trailing blank lines are dropped so the signature's own leading blank line
 * supplies the single gap above it.
 */
export function textToBlocks(text: string): string {
  const trimmed = text.replace(/\r/g, '').replace(/\s+$/, '');
  if (!trimmed) return '';
  const inline = (s: string) => linkifyEscapedText(escapeHtml(s));
  const style = withFontAttrs('');
  const out: string[] = [];
  // Asserted rather than annotated: flush() resets it inside a closure, which
  // control-flow narrowing cannot see, so an annotated `= null` stays `never`.
  let list = null as { tag: 'ul' | 'ol'; items: string[]; start: number } | null;
  const flush = () => {
    if (!list) return;
    const start = list.tag === 'ol' && list.start !== 1 ? ` start="${list.start}"` : '';
    const items = list.items.map((item) => `<li${style}>${item}</li>`).join('');
    out.push(`<${list.tag}${style}${start}>${items}</${list.tag}>`);
    list = null;
  };
  for (const line of trimmed.split('\n')) {
    const bullet = line.match(BULLET_ITEM);
    const numbered = bullet ? null : line.match(NUMBERED_ITEM);
    const tag = bullet ? 'ul' : numbered ? 'ol' : null;
    if (!tag) {
      flush();
      out.push(block(inline(line)));
      continue;
    }
    if (list?.tag !== tag) {
      flush();
      list = { tag, items: [], start: numbered ? Number(numbered[1]) : 1 };
    }
    list!.items.push(inline(bullet ? bullet[1] : numbered![2]));
  }
  flush();
  return out.join('');
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
 * Merge the outgoing font into one element's attributes, leaving an
 * author-specified font-family or font-size alone. `dropMargins` additionally
 * strips margin declarations, for the <p> rewrite that spaces blocks with a
 * blank line instead. Idempotent: a second pass finds the font already there.
 */
function withFontAttrs(attrs: string, opts?: { dropMargins?: boolean }): string {
  const styleMatch = attrs.match(/\bstyle\s*=\s*("|')([\s\S]*?)\1/i);
  if (!styleMatch) return `${attrs} style="${BLOCK_STYLE}"`;
  const decls = styleMatch[2]
    .split(';')
    .map((d) => d.trim())
    .filter((d) => d && !(opts?.dropMargins && /^margin/i.test(d)));
  if (!decls.some((d) => /^font-family/i.test(d))) decls.push(`font-family:${FONT_STACK}`);
  if (!decls.some((d) => /^font-size/i.test(d))) decls.push(`font-size:${FONT_SIZE}`);
  return attrs.replace(styleMatch[0], `style="${decls.join('; ')}"`);
}

/**
 * Rewrite authored <p> paragraphs into OWA-native blocks: each <p> becomes a
 * font-bearing <div> (see `block`), and the gap between paragraphs becomes a
 * blank-line block rather than a margin. Lists get the same blank line after
 * them that a paragraph does, unless they are nested inside another list item.
 *
 * <ul>/<ol>/<li> carry the font on the element itself rather than relying on
 * inheritance from the body wrapper, which is not always there to inherit
 * from: withOutgoingFont used to skip the wrapper for any body declaring a
 * font-family anywhere (every OWA-rewritten body does), leaving bare bullets
 * in OWA's base UI font beside Arial prose. Stamping the element covers that
 * path and anything else that loses the wrapper.
 *
 * It does NOT survive an edit in the OWA composer. OWA strips font-family and
 * font-size off <ul>/<li> on save — measured, not assumed: a draft stamped
 * here came back with the declarations deleted from the list elements while
 * identical ones on OWA's own block <div>s were kept. So the guarantee is at
 * send time and on the update_message path; a draft revised in the web
 * composer needs re-stamping before it goes out. Nothing server-side fixes
 * that, since OWA is normalizing its own stored HTML.
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
  let out = html.replace(
    /<p(\s[^>]*)?>/gi,
    (_match, attrs?: string) => `<div${withFontAttrs(attrs ?? '', { dropMargins: true })}>`,
  );
  out = out.replace(/<\/p>/gi, `</div>${BLANK_LINE}`);
  out = out.replace(
    /<(ul|ol|li)(\s[^>]*)?>/gi,
    (_match, tag: string, attrs?: string) => `<${tag}${withFontAttrs(attrs ?? '')}>`,
  );
  // The negative lookahead also covers a blank-line block that is already
  // there: without it every PATCH through withOutgoingFont added one more
  // blank line under each list, so a draft revised twice grew a widening gap.
  out = out.replace(
    /<\/(ul|ol)>(?!\s*(?:<\/(?:li|ul|ol)>|<div[^>]*>\s*<br\s*\/?>\s*<\/div>))/gi,
    (m) => m + BLANK_LINE,
  );
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
 * True when `html` is a single element that declares a font-family and
 * encloses the whole fragment, so that font already applies throughout. A
 * font-bearing element that covers only part of the fragment (one block among
 * siblings, as in an OWA-rewritten body) does not count.
 */
function isFontBearingRoot(html: string): boolean {
  const trimmed = html.trim();
  const open = trimmed.match(/^<([a-z][\w-]*)(?:\s[^>]*)?>/i);
  if (!open || !/font-family/i.test(open[0])) return false;
  const tags = new RegExp(`<${open[1]}(?:\\s[^>]*)?>|</${open[1]}\\s*>`, 'gi');
  let depth = 0;
  let match: RegExpExecArray | null;
  while ((match = tags.exec(trimmed)) !== null) {
    depth += match[0].startsWith('</') ? -1 : 1;
    // The root closes here; it is the whole fragment only if nothing follows.
    if (depth === 0) return tags.lastIndex === trimmed.length;
  }
  return false;
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
 * Idempotent: a body that is already one font-bearing root block — this
 * function's own wrapper, or an author's — is returned untouched rather than
 * nested inside a second wrapper. A full <html>…<body> document gets the
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

  // Wrap the body's own content: for a full <html> document that means inside
  // <body>, not around the document, which would be malformed.
  const bodyOpen = content.match(/<body[^>]*>/i);
  const hasBody = bodyOpen !== null && bodyOpen.index !== undefined;
  const start = hasBody ? bodyOpen!.index! + bodyOpen![0].length : 0;
  const closeAt = content.search(/<\/body>/i);
  const end = hasBody && closeAt > -1 ? closeAt : content.length;
  const inner = content.slice(start, end);

  // Already covered: nothing to wrap, or the content is a single font-bearing
  // root that encloses everything — our own wrapper from an earlier pass, or
  // an author's deliberate choice. Wrapping either again would nest one more
  // wrapper on every PATCH.
  //
  // This used to skip the wrapper whenever a font-family appeared ANYWHERE in
  // the body. That read as "respect the author", but it fired on every body
  // OWA had rewritten, because OWA stamps a font on each block <div> it owns.
  // Elements carrying no font of their own — <table>, <blockquote>, a bare
  // <div> — were then left inheriting OWA's base UI font, visibly mismatched
  // against the stamped blocks beside them. Wrapping those bodies is safe: an
  // inherited font loses to any inner declaration, so blocks that do carry
  // their own font are unaffected.
  if (!inner.trim() || isFontBearingRoot(inner)) return { contentType: 'html', content };

  const open = `<div style="${BODY_STYLE}">`;
  return {
    contentType: 'html',
    content: content.slice(0, start) + open + inner + '</div>' + content.slice(end),
  };
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
