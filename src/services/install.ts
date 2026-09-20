import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir, platform } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * MCP client configuration writer.
 *
 * Hey365 is a plain stdio MCP server, so it runs in any MCP-capable host.
 * The differences are only where the config lives and which envelope key it
 * uses, which is what this module encodes.
 */

export type ClientId =
  | 'copilot-cli'
  | 'vscode'
  | 'claude-code'
  | 'claude-desktop'
  | 'codex'
  | 'cursor'
  | 'windsurf'
  | 'scout';

export type ConfigFormat = 'mcpServers' | 'vscode-servers' | 'toml';

export interface ClientDefinition {
  id: ClientId;
  label: string;
  format: ConfigFormat;
  /** Resolves the config path; `cwd` is used by project-scoped clients. */
  path: (cwd: string) => string;
  scope: 'user' | 'project';
  note?: string;
}

const APP_DATA = process.env.APPDATA ?? join(homedir(), 'AppData', 'Roaming');

export const CLIENTS: ClientDefinition[] = [
  {
    id: 'copilot-cli',
    label: 'GitHub Copilot CLI',
    format: 'mcpServers',
    scope: 'user',
    path: () => join(homedir(), '.copilot', 'mcp-config.json'),
  },
  {
    id: 'vscode',
    label: 'GitHub Copilot (VS Code)',
    format: 'vscode-servers',
    scope: 'project',
    path: (cwd) => join(cwd, '.vscode', 'mcp.json'),
    note: 'ワークスペース単位の設定です。ユーザー全体に入れる場合は VS Code の MCP: Open User Configuration を使ってください。',
  },
  {
    id: 'claude-code',
    label: 'Claude Code',
    format: 'mcpServers',
    scope: 'project',
    path: (cwd) => join(cwd, '.mcp.json'),
    note: 'ユーザー全体に入れる場合は `claude mcp add --scope user hey365 -- node <dist>/index.js mcp`。',
  },
  {
    id: 'claude-desktop',
    label: 'Claude Desktop',
    format: 'mcpServers',
    scope: 'user',
    path: () =>
      platform() === 'darwin'
        ? join(homedir(), 'Library', 'Application Support', 'Claude', 'claude_desktop_config.json')
        : join(APP_DATA, 'Claude', 'claude_desktop_config.json'),
  },
  {
    id: 'codex',
    label: 'OpenAI Codex CLI',
    format: 'toml',
    scope: 'user',
    path: () => join(process.env.CODEX_HOME ?? join(homedir(), '.codex'), 'config.toml'),
  },
  {
    id: 'cursor',
    label: 'Cursor',
    format: 'mcpServers',
    scope: 'user',
    path: () => join(homedir(), '.cursor', 'mcp.json'),
  },
  {
    id: 'windsurf',
    label: 'Windsurf',
    format: 'mcpServers',
    scope: 'user',
    path: () => join(homedir(), '.codeium', 'windsurf', 'mcp_config.json'),
  },
  {
    id: 'scout',
    label: 'Scout',
    format: 'mcpServers',
    scope: 'user',
    path: () => join(homedir(), '.scout', 'mcp.json'),
    note: 'Scout の設定ファイル位置が異なる場合は --path で指定してください。',
  },
];

export interface ServerSpec {
  command: string;
  args: string[];
  env?: Record<string, string>;
}

/** Points the client at this installation's compiled entry point. */
export function buildServerSpec(options: { entry?: string; env?: Record<string, string> } = {}): ServerSpec {
  const entry = options.entry ?? defaultEntry();
  const spec: ServerSpec = { command: process.execPath, args: [entry, 'mcp'] };
  if (options.env && Object.keys(options.env).length > 0) spec.env = options.env;
  return spec;
}

export function defaultEntry(): string {
  // dist/services/install.js -> dist/index.js
  const here = dirname(fileURLToPath(import.meta.url));
  return resolve(here, '..', 'index.js');
}

export interface InstallOutcome {
  client: ClientId;
  label: string;
  path: string;
  written: boolean;
  created: boolean;
  preview: string;
  error?: string;
}

export function renderConfig(format: ConfigFormat, spec: ServerSpec, serverName = 'hey365'): string {
  if (format === 'toml') {
    const lines = [`[mcp_servers.${serverName}]`, `command = ${JSON.stringify(spec.command)}`];
    lines.push(`args = [${spec.args.map((arg) => JSON.stringify(arg)).join(', ')}]`);
    if (spec.env) {
      const entries = Object.entries(spec.env).map(([key, value]) => `${key} = ${JSON.stringify(value)}`);
      lines.push(`env = { ${entries.join(', ')} }`);
    }
    return `${lines.join('\n')}\n`;
  }

  const envelope =
    format === 'vscode-servers'
      ? { servers: { [serverName]: { type: 'stdio', ...spec } } }
      : { mcpServers: { [serverName]: spec } };
  return `${JSON.stringify(envelope, null, 2)}\n`;
}

export function installForClient(
  client: ClientDefinition,
  spec: ServerSpec,
  options: { cwd?: string; dryRun?: boolean; pathOverride?: string; serverName?: string } = {},
): InstallOutcome {
  const serverName = options.serverName ?? 'hey365';
  const target = options.pathOverride ?? client.path(options.cwd ?? process.cwd());
  const preview = renderConfig(client.format, spec, serverName);

  const outcome: InstallOutcome = {
    client: client.id,
    label: client.label,
    path: target,
    written: false,
    created: false,
    preview,
  };

  if (options.dryRun) return outcome;

  try {
    let existing = '';
    let created = true;
    try {
      existing = readFileSync(target, 'utf8');
      created = false;
    } catch {
      existing = '';
    }

    const merged =
      client.format === 'toml'
        ? mergeToml(existing, serverName, spec)
        : mergeJson(existing, client.format, serverName, spec);

    mkdirSync(dirname(target), { recursive: true });
    writeFileSync(target, merged, 'utf8');
    outcome.written = true;
    outcome.created = created;
  } catch (error) {
    outcome.error = (error as Error).message;
  }

  return outcome;
}

/** Adds the server entry without discarding the user's other servers. */
export function mergeJson(existing: string, format: ConfigFormat, serverName: string, spec: ServerSpec): string {
  const key = format === 'vscode-servers' ? 'servers' : 'mcpServers';
  let root: Record<string, unknown> = {};
  if (existing.trim()) {
    try {
      root = JSON.parse(existing) as Record<string, unknown>;
    } catch {
      throw new Error(`既存の設定ファイルが JSON として解析できません: ${serverName}`);
    }
  }
  const bucket = (root[key] && typeof root[key] === 'object' ? root[key] : {}) as Record<string, unknown>;
  bucket[serverName] = format === 'vscode-servers' ? { type: 'stdio', ...spec } : { ...spec };
  root[key] = bucket;
  return `${JSON.stringify(root, null, 2)}\n`;
}

/** Replaces an existing `[mcp_servers.<name>]` table or appends a new one. */
export function mergeToml(existing: string, serverName: string, spec: ServerSpec): string {
  const block = renderConfig('toml', spec, serverName);
  if (!existing.trim()) return block;

  const header = new RegExp(`^\\[mcp_servers\\.${serverName}\\]\\s*$`, 'm');
  const match = header.exec(existing);
  if (!match || match.index === undefined) {
    const separator = existing.endsWith('\n') ? '\n' : '\n\n';
    return `${existing}${separator}${block}`;
  }

  const after = existing.slice(match.index + match[0].length);
  const nextTable = /^\s*\[/m.exec(after);
  const end = nextTable?.index === undefined ? existing.length : match.index + match[0].length + nextTable.index;
  return `${existing.slice(0, match.index)}${block}${existing.slice(end)}`;
}

export function findClient(id: string): ClientDefinition | undefined {
  return CLIENTS.find((client) => client.id === id);
}
