import {
  SIGNATURE_HTML,
  SIGNATURE_TEXT,
  withSignature,
  buildReplyTopHtml,
  injectReplyTop,
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
