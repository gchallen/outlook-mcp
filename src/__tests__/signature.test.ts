import {
  SIGNATURE_HTML,
  SIGNATURE_TEXT,
  withSignature,
  buildReplyTopHtml,
  injectReplyTop,
  normalizeParagraphs,
  block,
  BLANK_LINE,
  withOutgoingFont,
  linkifyEscapedText,
  linkifyHtml,
  textToHtml,
} from '../signature.js';

describe('signature', () => {
  describe('withSignature', () => {
    it('returns undefined when body is undefined', () => {
      expect(withSignature(undefined)).toBeUndefined();
    });

    it('wraps an HTML body in the font stack and appends the HTML signature', () => {
      const result = withSignature({ contentType: 'html', content: '<div>Hello</div>' });
      expect(result).toEqual({
        contentType: 'html',
        content:
          '<div style="font-family:Arial,Helvetica,sans-serif; font-size:11pt"><div>Hello</div></div>' +
          SIGNATURE_HTML,
      });
    });

    it('converts a text body to HTML with the styled signature', () => {
      const result = withSignature({ contentType: 'text', content: 'Hello' });
      expect(result?.contentType).toBe('html');
      expect(result?.content).toContain(block('Hello'));
      expect(result?.content).toContain(SIGNATURE_HTML);
    });

    it('escapes markup and hyperlinks bare URLs in a text body', () => {
      const result = withSignature({
        contentType: 'text',
        content: 'See <this>: https://example.com/page',
      });
      expect(result?.content).toContain('See &lt;this&gt;:');
      expect(result?.content).toContain(
        '<a href="https://example.com/page" style="color:rgb(12,100,192)">https://example.com/page</a>',
      );
    });

    it('replaces a spelled-out plain signature with the styled one', () => {
      const result = withSignature({
        contentType: 'text',
        content: `Thanks!\n\n${SIGNATURE_TEXT}`,
      });
      expect(result?.content).toContain('Thanks!');
      expect(result?.content).toContain(SIGNATURE_HTML);
      // The plain copy is gone: the only geoffreychallen.com mentions are the
      // signature's own href/title/label.
      expect(result?.content!.match(/geoffreychallen\.com/g)).toHaveLength(3);
    });

    it('treats a missing contentType as text and still emits HTML', () => {
      const result = withSignature({ content: 'Hi' });
      expect(result?.contentType).toBe('html');
      expect(result?.content).toContain(block('Hi'));
      expect(result?.content).toContain(SIGNATURE_HTML);
    });

    it('uses just the signature when text content is empty', () => {
      const result = withSignature({ contentType: 'text', content: '' });
      expect(result).toEqual({ contentType: 'html', content: SIGNATURE_HTML });
    });

    it('is case-insensitive about the HTML content type', () => {
      const result = withSignature({ contentType: 'HTML', content: '<p>x</p>' });
      expect(result?.contentType).toBe('html');
      expect(result?.content).toContain(SIGNATURE_HTML);
    });

    it('hyperlinks bare URLs in an HTML body without touching existing links', () => {
      const result = withSignature({
        contentType: 'html',
        content: '<p>See https://example.com and <a href="https://x.com">x</a></p>',
      });
      expect(result?.content).toContain(
        '<a href="https://example.com" style="color:rgb(12,100,192)">https://example.com</a>',
      );
      expect(result?.content).toContain('<a href="https://x.com">x</a>');
    });

    it('does not double-sign a body that already contains the signature', () => {
      const signed = { contentType: 'html', content: '<div>Hi</div>' + SIGNATURE_HTML };
      expect(withSignature(signed)).toBe(signed);
    });

    it('still signs a text body that merely mentions the signature url', () => {
      const result = withSignature({ contentType: 'text', content: 'see geoffreychallen.com' });
      expect(result?.contentType).toBe('html');
      expect(result?.content).toContain('see geoffreychallen.com');
      expect(result?.content).toContain(SIGNATURE_HTML);
    });
  });

  describe('SIGNATURE_HTML', () => {
    it('dims the name with opacity instead of a hardcoded gray', () => {
      expect(SIGNATURE_HTML).toContain('opacity:0.6');
      expect(SIGNATURE_HTML).not.toContain('rgb(102,102,102)');
      expect(SIGNATURE_HTML).not.toContain('rgb(51,51,51)');
    });

    it('keeps an explicit blue link affordance', () => {
      expect(SIGNATURE_HTML).toContain('color:rgb(12,100,192)');
      expect(SIGNATURE_HTML).toContain('href="https://geoffreychallen.com"');
    });
  });

  describe('buildReplyTopHtml', () => {
    it('includes only the signature when there is no comment', () => {
      expect(buildReplyTopHtml()).toBe(SIGNATURE_HTML);
    });

    it('prepends an escaped comment before the signature', () => {
      const html = buildReplyTopHtml('Thanks <Jim> & co');
      expect(html).toContain('Thanks &lt;Jim&gt; &amp; co');
      expect(html.endsWith(SIGNATURE_HTML)).toBe(true);
    });

    it('renders each comment line as an OWA-style block', () => {
      const html = buildReplyTopHtml('line1\nline2');
      expect(html).toContain(block('line1') + block('line2'));
    });

    it('renders a blank line in the comment as an empty block, like OWA', () => {
      const html = buildReplyTopHtml('one\n\ntwo');
      expect(html).toContain(block('one') + BLANK_LINE + block('two'));
    });

    it('lets the signature supply the single blank line above itself', () => {
      const html = buildReplyTopHtml('line1\nline2\n\n');
      expect(html).toContain(`${block('line2')}</div>${SIGNATURE_HTML}`);
      expect(SIGNATURE_HTML.startsWith(`<div id="Signature">${BLANK_LINE}`)).toBe(true);
    });

    it('does not stack a second blank line between a raw HTML comment and the signature', () => {
      const html = buildReplyTopHtml('<p>One</p>', { raw: true });
      expect(html).toContain(`${block('One')}</div>${SIGNATURE_HTML}`);
      expect(html).not.toContain(BLANK_LINE + BLANK_LINE);
    });

    it('replaces a spelled-out plain signature with the styled one', () => {
      const html = buildReplyTopHtml(`Thanks!\n\n${SIGNATURE_TEXT}`);
      expect(html).toContain('Thanks!');
      expect(html.endsWith(SIGNATURE_HTML)).toBe(true);
      // Only the styled signature mentions the domain (href/title/label).
      expect(html.match(/geoffreychallen\.com/g)).toHaveLength(3);
    });

    it('hyperlinks bare URLs in a plain-text comment', () => {
      const html = buildReplyTopHtml(
        'Try https://www.usingandunderstanding.ai/educators#conversational-assessment today',
      );
      expect(html).toContain(
        '<a href="https://www.usingandunderstanding.ai/educators#conversational-assessment" style="color:rgb(12,100,192)">',
      );
    });

    it('leaves trailing punctuation outside a hyperlinked URL', () => {
      const html = buildReplyTopHtml('See https://example.com/a.');
      expect(html).toContain(
        '<a href="https://example.com/a" style="color:rgb(12,100,192)">https://example.com/a</a>.',
      );
    });

    it('hyperlinks bare URLs in a raw HTML comment without double-linking', () => {
      const html = buildReplyTopHtml(
        '<p>Bare https://example.com and linked <a href="https://x.com">https://x.com</a></p>',
        { raw: true },
      );
      expect(html).toContain(
        '<a href="https://example.com" style="color:rgb(12,100,192)">https://example.com</a>',
      );
      expect(html).toContain('<a href="https://x.com">https://x.com</a>');
    });

    it('does not append a second signature when a raw comment already carries the styled one', () => {
      const html = buildReplyTopHtml(`<p>Thanks!</p>${SIGNATURE_HTML}`, { raw: true });
      expect(html.match(/id="Signature"/g)).toHaveLength(1);
    });

    it('rewrites raw HTML <p> blocks as OWA-native blocks separated by a blank line', () => {
      const html = buildReplyTopHtml('<p>One</p><p>Two</p>', { raw: true });
      expect(html).toContain(block('One') + BLANK_LINE + block('Two'));
      expect(html).not.toContain('<p');
    });

    it('renders the comment OWA-native: no max-width, no opacity, no hardcoded color', () => {
      const html = buildReplyTopHtml('Hi');
      // Inspect just the comment wrapper, above the (separately styled) signature.
      const commentPart = html.slice(0, html.indexOf(SIGNATURE_HTML));
      // Earlier versions narrowed and dimmed authored comments, which made them
      // visibly different from paragraphs typed in the OWA composer.
      expect(commentPart).not.toContain('max-width');
      expect(commentPart).not.toContain('opacity');
      // No fixed color that would break the recipient's dark mode.
      expect(commentPart).not.toContain('color:');
    });
  });

  describe('normalizeParagraphs', () => {
    const BLOCK = 'font-family:Arial,Helvetica,sans-serif; font-size:11pt';

    it('turns a bare <p> into a font-bearing <div> with no trailing blank line', () => {
      expect(normalizeParagraphs('<p>Hi</p>')).toBe(`<div style="${BLOCK}">Hi</div>`);
    });

    it('preserves other attributes while converting', () => {
      expect(normalizeParagraphs('<p class="x">Hi</p>')).toBe(
        `<div class="x" style="${BLOCK}">Hi</div>`,
      );
    });

    it('keeps an author style but adds the missing font declarations', () => {
      expect(normalizeParagraphs('<p style="color:red">Hi</p>')).toBe(
        `<div style="color:red; ${BLOCK}">Hi</div>`,
      );
    });

    it("drops author margins: spacing is the blank line's job now", () => {
      expect(normalizeParagraphs('<p style="margin:2em 0">Hi</p>')).toBe(
        `<div style="${BLOCK}">Hi</div>`,
      );
    });

    it('respects an author font', () => {
      expect(normalizeParagraphs('<p style="font-family:Georgia">Hi</p>')).toBe(
        '<div style="font-family:Georgia; font-size:11pt">Hi</div>',
      );
    });

    it('is case-insensitive and separates paragraphs with a blank line', () => {
      expect(normalizeParagraphs('<P>a</P><p>b</p>')).toBe(
        `<div style="${BLOCK}">a</div>${BLANK_LINE}<div style="${BLOCK}">b</div>`,
      );
    });

    it('does not touch non-paragraph tags', () => {
      const html = '<ul><li>a</li></ul>';
      expect(normalizeParagraphs(html)).toBe(html);
    });

    it('gives a list the same blank line after it that a paragraph gets', () => {
      expect(normalizeParagraphs('<p>a</p><ul><li>b</li></ul><p>c</p>')).toBe(
        `${block('a')}${BLANK_LINE}<ul><li>b</li></ul>${BLANK_LINE}${block('c')}`,
      );
    });

    it('does not insert a blank line inside a nested list', () => {
      const html = '<ul><li>a<ul><li>b</li></ul></li></ul>';
      expect(normalizeParagraphs(html)).toBe(html);
    });

    it('removes the blank line before an existing signature, which carries its own', () => {
      expect(normalizeParagraphs(`<p>Hi</p>${SIGNATURE_HTML}`)).toBe(block('Hi') + SIGNATURE_HTML);
    });
  });

  describe('injectReplyTop', () => {
    it('inserts the comment and signature just after the <body> tag', () => {
      const draft = {
        contentType: 'html',
        content: '<html><body dir="ltr"><div id="divRplyFwdMsg">quoted</div></body></html>',
      };
      const result = injectReplyTop(draft, 'My reply');

      expect(result.contentType).toBe('html');
      // Signature lands above the quoted thread.
      const sigIndex = result.content!.indexOf('geoffreychallen.com');
      const quotedIndex = result.content!.indexOf('divRplyFwdMsg');
      expect(sigIndex).toBeGreaterThan(-1);
      expect(sigIndex).toBeLessThan(quotedIndex);
      expect(result.content).toContain('My reply');
      // Original markup is preserved.
      expect(result.content).toContain('<body dir="ltr">');
      expect(result.content).toContain('quoted');
    });

    it('prepends content when there is no <body> tag', () => {
      const draft = { contentType: 'html', content: '<div>quoted</div>' };
      const result = injectReplyTop(draft, 'Hi');
      expect(result.content!.indexOf('geoffreychallen.com')).toBeLessThan(
        result.content!.indexOf('quoted'),
      );
    });

    it('converts a text draft skeleton to signed HTML', () => {
      const draft = { contentType: 'text', content: '> quoted line' };
      const result = injectReplyTop(draft, 'My reply');
      expect(result.contentType).toBe('html');
      expect(result.content).toContain('My reply');
      expect(result.content).toContain(SIGNATURE_HTML);
      expect(result.content).toContain('&gt; quoted line');
      // New content sits above the quoted history.
      expect(result.content!.indexOf('My reply')).toBeLessThan(
        result.content!.indexOf('&gt; quoted line'),
      );
    });

    it('signs even without a comment', () => {
      const draft = { contentType: 'html', content: '<html><body></body></html>' };
      const result = injectReplyTop(draft);
      expect(result.content).toContain('geoffreychallen.com');
    });

    it('defaults an undefined draft body to empty html', () => {
      const result = injectReplyTop(undefined, 'hi');
      expect(result.contentType).toBe('html');
      expect(result.content).toContain('geoffreychallen.com');
    });
  });
});

describe('withOutgoingFont', () => {
  it('wraps a bare HTML body in the outgoing font', () => {
    const out = withOutgoingFont({ contentType: 'html', content: '<p>Hi</p>' });
    expect(out?.content).toContain('font-family:Arial,Helvetica,sans-serif');
    expect(out?.content).toContain('font-size:11pt');
  });

  it('rewrites bare paragraphs as blocks separated by a blank line', () => {
    const out = withOutgoingFont({ contentType: 'html', content: '<p>One</p><p>Two</p>' });
    expect(out?.content).toContain(block('One') + BLANK_LINE + block('Two'));
    expect(out?.content).not.toContain('margin');
  });

  it('separates a plain-text replacement body from its upgraded signature with one blank line', () => {
    const out = withOutgoingFont({ contentType: 'text', content: `Hi\n\n${SIGNATURE_TEXT}` });
    expect(out?.content).toContain(`${block('Hi')}</div>${SIGNATURE_HTML}`);
  });

  it('leaves an author-styled body alone', () => {
    const content = '<p style="font-family:Georgia">Hi</p>';
    const out = withOutgoingFont({ contentType: 'html', content });
    expect(out?.content).toContain('Georgia');
    expect(out?.content).not.toContain('Arial,Helvetica');
  });

  it('wraps inside <body> for a full document rather than around it', () => {
    const content = '<html><head></head><body dir="ltr"><p>Hi</p></body></html>';
    const out = withOutgoingFont({ contentType: 'html', content });
    expect(out?.content).toMatch(/<body dir="ltr"><div style="font-family:Arial/);
    expect(out?.content).toMatch(/<\/div><\/body>/);
  });

  it('converts a plain-text replacement body to HTML with links', () => {
    const out = withOutgoingFont({ contentType: 'text', content: 'Hi https://example.com' });
    expect(out?.contentType).toBe('html');
    expect(out?.content).toContain('font-family:Arial,Helvetica,sans-serif');
    expect(out?.content).toContain(
      '<a href="https://example.com" style="color:rgb(12,100,192)">https://example.com</a>',
    );
    // No signature added: on the update path it is the caller's to include.
    expect(out?.content).not.toContain('id="Signature"');
  });

  it('upgrades a plain-text signature in a replacement body to the styled one', () => {
    const out = withOutgoingFont({ contentType: 'text', content: `Hi\n\n${SIGNATURE_TEXT}` });
    expect(out?.content).toContain(SIGNATURE_HTML);
    expect(out?.content!.match(/geoffreychallen\.com/g)).toHaveLength(3);
  });

  it('hyperlinks bare URLs in an HTML replacement body', () => {
    const out = withOutgoingFont({
      contentType: 'html',
      content: '<p>See https://example.com</p>',
    });
    expect(out?.content).toContain(
      '<a href="https://example.com" style="color:rgb(12,100,192)">https://example.com</a>',
    );
  });
});

describe('linkify', () => {
  it('hyperlinks a URL in escaped text and keeps entity escapes intact', () => {
    expect(linkifyEscapedText('go to https://example.com/a?b=1&amp;c=2 now')).toBe(
      'go to <a href="https://example.com/a?b=1&amp;c=2" style="color:rgb(12,100,192)">https://example.com/a?b=1&amp;c=2</a> now',
    );
  });

  it('leaves text without URLs unchanged', () => {
    expect(linkifyEscapedText('no links here')).toBe('no links here');
  });

  it('does not linkify inside href attributes or existing anchors', () => {
    const html = '<a href="https://x.com" title="https://x.com">click https://x.com</a>';
    expect(linkifyHtml(html)).toBe(html);
  });

  it('does not linkify inside <style> blocks', () => {
    const html = '<style>p { background:url(https://x.com/i.png) }</style>';
    expect(linkifyHtml(html)).toBe(html);
  });

  it('textToHtml escapes, links, and converts newlines', () => {
    expect(textToHtml('a & b\nhttps://x.com')).toBe(
      'a &amp; b<br><a href="https://x.com" style="color:rgb(12,100,192)">https://x.com</a>',
    );
  });
});
