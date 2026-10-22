import { execFileSync } from 'child_process';
import { readFileSync } from 'fs';
import { join } from 'path';
import { GUIDE_URL, PACKAGE_NAME } from '../config';
import { buildTools } from '../server';

// What ships: the npm package, the Claude Desktop extension manifest, and the
// docs people follow. These checks catch drift between them, like a bin entry
// npm silently drops or a guide that names a tool that no longer exists.

const ROOT = join(__dirname, '../..');
const read = (file: string) => readFileSync(join(ROOT, file), 'utf-8');
const pkg = JSON.parse(read('package.json'));
const manifest = JSON.parse(read('manifest.json'));

describe('npm package', () => {
  let files: string[];

  beforeAll(() => {
    execFileSync('npx', ['tsc'], { cwd: ROOT, stdio: 'pipe' });
    const output = execFileSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], {
      cwd: ROOT,
      encoding: 'utf-8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    // npm 10 prints an array of packages; npm 11+ an object keyed by name.
    const parsed = JSON.parse(output);
    const packed = Array.isArray(parsed) ? parsed[0] : parsed[pkg.name];
    files = packed.files.map((file: { path: string }) => file.path);
  }, 120_000);

  it('is the published name the docs use', () => {
    expect(pkg.name).toBe(PACKAGE_NAME);
    expect(pkg.publishConfig).toEqual({ access: 'public' });
  });

  it('uses date-based versioning', () => {
    expect(pkg.version).toMatch(/^20\d\d\.(1[0-2]|[1-9])\.\d+$/);
  });

  it('declares a bin npm accepts, pointing at a shipped file', () => {
    // npm drops a bin path written as "./dist/cli.js" when publishing.
    expect(pkg.bin).toEqual({ 'outlook-mcp': 'dist/cli.js' });
    expect(files).toContain('dist/cli.js');
  });

  it('ships the server, setup, and docs', () => {
    for (const file of [
      'dist/cli.js',
      'dist/index.js',
      'dist/server.js',
      'dist/setup.js',
      'dist/login.js',
      'docs/GUIDED-SETUP.md',
      'manifest.json',
      'README.md',
    ]) {
      expect(files).toContain(file);
    }
  });

  it('ships no sources, tests, or secrets', () => {
    for (const file of files) {
      expect(file).not.toMatch(/^src\/|__tests__|\.test\.|\.env|token-cache|\.mcpb$|^coverage\//);
    }
  });

  it('lists every runtime import as a dependency', () => {
    const imported = new Set<string>();
    for (const file of files.filter((name) => name.endsWith('.js'))) {
      for (const [, from, bare] of read(file).matchAll(
        /from '([^'.][^']*)'|import '([^'.][^']*)'/g,
      )) {
        const spec = from ?? bare;
        const name = spec.startsWith('@')
          ? spec.split('/').slice(0, 2).join('/')
          : spec.split('/')[0];
        imported.add(name);
      }
    }
    expect(imported).toContain('isomorphic-fetch');
    expect(imported).toContain('@azure/msal-node');
    const builtins = new Set(['fs', 'os', 'path', 'readline', 'child_process', 'crypto', 'url']);
    for (const name of imported) {
      if (builtins.has(name) || name.startsWith('node:')) continue;
      expect({ name, listed: name in pkg.dependencies }).toEqual({ name, listed: true });
    }
  });
});

describe('Claude Desktop extension manifest', () => {
  it('matches the package version', () => {
    expect(manifest.version).toBe(pkg.version);
  });

  it('runs the shipped CLI', () => {
    expect(manifest.server.entry_point).toBe(pkg.bin['outlook-mcp']);
    expect(manifest.server.mcp_config.args).toEqual([`\${__dirname}/${pkg.bin['outlook-mcp']}`]);
  });

  it('only passes settings the server reads, from fields it defines', () => {
    const config = read('src/config.ts');
    for (const [name, value] of Object.entries<string>(manifest.server.mcp_config.env)) {
      expect(config).toContain(`'${name}'`);
      const field = value.match(/^\$\{user_config\.([a-z_]+)\}$/)?.[1];
      expect({ name, field: field && field in manifest.user_config }).toEqual({
        name,
        field: true,
      });
    }
  });

  it('requires only the email and the client ID', () => {
    const required = Object.entries<{ required?: boolean }>(manifest.user_config)
      .filter(([, field]) => field.required)
      .map(([key]) => key);
    expect(required.sort()).toEqual(['client_id', 'email']);
  });

  it('links to the guided setup', () => {
    expect(manifest.documentation).toBe(GUIDE_URL);
  });
});

describe('docs', () => {
  const toolNames = new Set(buildTools().map((tool) => tool.name));

  it.each(['README.md', 'docs/GUIDED-SETUP.md', 'SETUP.md', 'INSTRUCTIONS.md'])(
    '%s names only tools that exist',
    (file) => {
      const text = read(file);
      const mentioned = [
        ...text.matchAll(/`([a-z]+(?:_[a-z]+)+)`/g),
        ...text.matchAll(/\b(outlook_status|[gs]et_signature)\b/g),
      ].map((match) => match[1]);
      const notTools = new Set(['client_id', 'user_config', 'extra_scopes', 'mcp_servers']);
      for (const name of mentioned) {
        if (notTools.has(name) || /^[a-z]+_[a-z]+_?(id|home|cache|scopes)$/.test(name)) continue;
        if (
          !/^(list|get|set|create|update|delete|send|reply|forward|move|copy|search|save|add|mark|react|respond|outlook)_/.test(
            name,
          )
        )
          continue;
        expect({ file, name, exists: toolNames.has(name) }).toEqual({ file, name, exists: true });
      }
    },
  );

  it.each(['README.md', 'docs/GUIDED-SETUP.md', 'SETUP.md', 'INSTRUCTIONS.md'])(
    '%s uses the published package name',
    (file) => {
      for (const [, name] of read(file).matchAll(/npx -y ([^\s`]+)/g)) {
        expect(name).toBe(PACKAGE_NAME);
      }
    },
  );

  it('guided setup links to the release with the extension', () => {
    expect(read('docs/GUIDED-SETUP.md')).toContain(
      'https://github.com/gchallen/outlook-mcp/releases/latest',
    );
  });

  it('contain no real tenant or client IDs', () => {
    const allowed = new Set([
      '11111111-2222-4333-8444-555555555555',
      'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
      '9188040d-6c67-4c5b-b112-36a304b66dad',
      '1a2b3c4d-1234-5678-9abc-1234567890ab',
    ]);
    for (const file of [
      'README.md',
      'docs/GUIDED-SETUP.md',
      'SETUP.md',
      'INSTRUCTIONS.md',
      '.env.example',
    ]) {
      for (const [guid] of read(file).matchAll(/[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}/gi)) {
        expect({ file, guid, allowed: allowed.has(guid.toLowerCase()) }).toEqual({
          file,
          guid,
          allowed: true,
        });
      }
    }
  });
});
