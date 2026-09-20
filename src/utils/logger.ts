/**
 * stderr-only logger with secret redaction.
 *
 * MCP stdio servers must never write to stdout outside the protocol stream,
 * so every diagnostic goes to stderr. Section 18 of the spec forbids logging
 * credentials, so all output passes through `redact()` first.
 */

export type LogLevel = 'silent' | 'error' | 'warn' | 'info' | 'debug';

const LEVEL_ORDER: Record<LogLevel, number> = {
  silent: 0,
  error: 1,
  warn: 2,
  info: 3,
  debug: 4,
};

function currentLevel(): LogLevel {
  const raw = (process.env.HEY365_LOG_LEVEL ?? 'info').toLowerCase();
  return (raw in LEVEL_ORDER ? raw : 'info') as LogLevel;
}

const SECRET_PATTERNS: Array<[RegExp, string]> = [
  // Bearer / JWT style tokens.
  [/\bey[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b/g, '[REDACTED_JWT]'],
  [/\bbearer\s+[A-Za-z0-9._~+/-]{10,}=*/gi, 'Bearer [REDACTED]'],
  // Structured secret fields in JSON or key=value form.
  [
    /("?(?:access_?token|refresh_?token|id_?token|client_?secret|password|pwd|secret|authorization|cookie|set-cookie|api[_-]?key)"?\s*[:=]\s*)("[^"]*"|'[^']*'|[^\s,;}]+)/gi,
    '$1"[REDACTED]"',
  ],
  // Long opaque MSAL-ish blobs.
  [/\b[A-Za-z0-9_-]{5,}~[A-Za-z0-9._~-]{20,}\b/g, '[REDACTED_SECRET]'],
];

export function redact(input: string): string {
  let output = input;
  for (const [pattern, replacement] of SECRET_PATTERNS) {
    output = output.replace(pattern, replacement);
  }
  return output;
}

function write(level: Exclude<LogLevel, 'silent'>, message: string, meta?: unknown): void {
  if (LEVEL_ORDER[currentLevel()] < LEVEL_ORDER[level]) return;
  const stamp = new Date().toISOString();
  let line = `[hey365] ${stamp} ${level.toUpperCase()} ${message}`;
  if (meta !== undefined) {
    let serialised: string;
    try {
      serialised = typeof meta === 'string' ? meta : JSON.stringify(meta);
    } catch {
      serialised = String(meta);
    }
    line += ` ${serialised}`;
  }
  process.stderr.write(`${redact(line)}\n`);
}

export const logger = {
  error: (message: string, meta?: unknown) => write('error', message, meta),
  warn: (message: string, meta?: unknown) => write('warn', message, meta),
  info: (message: string, meta?: unknown) => write('info', message, meta),
  debug: (message: string, meta?: unknown) => write('debug', message, meta),
};
