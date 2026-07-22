import {
  SIGNATURE_HTML,
  SIGNATURE_TEXT,
  withSignature,
  buildReplyTopHtml,
  injectReplyTop,
  ensureParagraphSpacing,
} from '../signature';

describe('signature', () => {
  describe('withSignature', () => {
    it('returns undefined when body is undefined', () => {
      expect(withSignature(undefined)).toBeUndefined();
    });

    it('appends the HTML signature to an HTML body', () => {
      const result = withSignature({ contentType: 'html', content: '<div>Hello</div>' });
      expect(result).toEqual({
        contentType: 'html',
        content: '<div>Hello</div>' + SIGNATURE_HTML,
      });
    });

    it('appends the text signature to a text body', () => {
      const result = withSignature({ contentType: 'text', content: 'Hello' });
      expect(result).toEqual({ contentType: 'text', content: `Hello\n\n${SIGNATURE_TEXT}` });
    });

    it('treats a missing contentType as text', () => {
      const result = withSignature({ content: 'Hi' });
      expect(result).toEqual({ contentType: 'text', content: `Hi\n\n${SIGNATURE_TEXT}` });
    });

    it('uses just the signature when text content is empty', () => {
      const result = withSignature({ contentType: 'text', content: '' });
      expect(result).toEqual({ contentType: 'text', content: SIGNATURE_TEXT });
    });

    it('is case-insensitive about the HTML content type', () => {
      const result = withSignature({ contentType: 'HTML', content: '<p>x</p>' });
      expect(result?.contentType).toBe('html');
      expect(result?.content).toContain(SIGNATURE_HTML);
    });

    it('does not double-sign a body that already contains the signature', () => {
      const signed = { contentType: 'html', content: '<div>Hi</div>' + SIGNATURE_HTML };
      expect(withSignature(signed)).toBe(signed);
    });

    it('does not double-sign a text body that mentions the signature url', () => {
      const body = { contentType: 'text', content: 'see geoffreychallen.com' };
      expect(withSignature(body)).toBe(body);
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

    it('converts newlines in the comment to <br>', () => {
      const html = buildReplyTopHtml('line1\nline2');
      expect(html).toContain('line1<br>line2');
    });

    it('does not append a second signature when the comment is already signed', () => {
      const html = buildReplyTopHtml(`Thanks!\n\n${SIGNATURE_TEXT}`);
      expect(html).not.toContain(SIGNATURE_HTML);
      expect(html.match(/geoffreychallen\.com/g)).toHaveLength(1);
    });

    it('styles raw HTML <p> blocks so they survive the margin reset with even leading', () => {
      const html = buildReplyTopHtml('<p>One</p><p>Two</p>', { raw: true });
      expect(html).toContain('<p style="margin:0 0 1em 0; font-size:11pt; line-height:1.4">One</p>');
      expect(html).toContain('<p style="margin:0 0 1em 0; font-size:11pt; line-height:1.4">Two</p>');
    });

    it('wraps the comment with max-width and opacity but no hardcoded text color', () => {
      const html = buildReplyTopHtml('Hi');
      // Inspect just the comment wrapper, above the (separately styled) signature.
      const commentPart = html.slice(0, html.indexOf(SIGNATURE_HTML));
      expect(commentPart).toContain('max-width:40em');
      expect(commentPart).toContain('opacity:0.8');
      // No fixed color that would break the recipient's dark mode.
      expect(commentPart).not.toContain('color:');
    });
  });

  describe('ensureParagraphSpacing', () => {
    const BLOCK = 'margin:0 0 1em 0; font-size:11pt; line-height:1.4';

    it('gives a bare <p> the block style (spacing + font-size + line-height)', () => {
      expect(ensureParagraphSpacing('<p>Hi</p>')).toBe(`<p style="${BLOCK}">Hi</p>`);
    });

    it('preserves other attributes while adding the block style', () => {
      expect(ensureParagraphSpacing('<p class="x">Hi</p>')).toBe(
        `<p class="x" style="${BLOCK}">Hi</p>`,
      );
    });

    it('merges a margin into an existing style that lacks one', () => {
      expect(ensureParagraphSpacing('<p style="color:red">Hi</p>')).toBe(
        '<p style="color:red; margin:0 0 1em 0">Hi</p>',
      );
    });

    it('leaves an author-specified margin untouched', () => {
      const html = '<p style="margin:2em 0">Hi</p>';
      expect(ensureParagraphSpacing(html)).toBe(html);
    });

    it('is case-insensitive and handles multiple paragraphs', () => {
      expect(ensureParagraphSpacing('<P>a</P><p>b</p>')).toBe(
        `<P style="${BLOCK}">a</P><p style="${BLOCK}">b</p>`,
      );
    });

    it('does not touch non-paragraph tags', () => {
      const html = '<ul><li>a</li></ul>';
      expect(ensureParagraphSpacing(html)).toBe(html);
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

    it('handles a text draft body', () => {
      const draft = { contentType: 'text', content: '> quoted line' };
      const result = injectReplyTop(draft, 'My reply');
      expect(result).toEqual({
        contentType: 'text',
        content: `My reply\n\n${SIGNATURE_TEXT}\n\n> quoted line`,
      });
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
