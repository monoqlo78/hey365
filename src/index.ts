#!/usr/bin/env node
import { formatSessionSummary } from './services/format.js';
import { CLIENTS, buildServerSpec, findClient, installForClient, renderConfig } from './services/install.js';
import { summarizeSession } from './services/session-summary.js';
import { checkHealth, runSetup } from './services/setup.js';
import { startStdioServer, HEY365_VERSION } from './server.js';
import { triageTool } from './tools/index.js';
import { asHey365Error } from './utils/errors.js';
import { timezone } from './utils/time.js';

const USAGE = `hey365 ${HEY365_VERSION}

Usage:
  hey365 mcp                         MCP stdio サーバーを起動する（既定）
  hey365 triage [--hours 36]         ターミナルからトリアージを実行する
  hey365 session <sessionId>         会議 / 会話を要約する
  hey365 health [--deep]             接続状態を確認する
  hey365 setup [--no-interactive]    Work IQ のインストール / 認証を復旧する
  hey365 install [client...]         MCP クライアントに Hey365 を登録する
  hey365 --help

Install targets:
${CLIENTS.map((client) => `  ${client.id.padEnd(15)} ${client.label}`).join('\n')}

Options:
  --hours <n>        triage の対象時間（既定 36）
  --limit <n>        triage の最大件数（既定 10）
  --no-drafts        返信案を生成しない
  --deep             health で各サービスの到達性も確認する
  --dry-run          install で書き込まずに内容だけ表示する
  --path <file>      install の書き込み先を上書きする
  --all              install で対応クライアント全てに書き込む

Environment:
  HEY365_TIMEZONE          表示タイムゾーン（既定 Asia/Tokyo）
  HEY365_VIP               重要送信者のメールアドレス（カンマ区切り）
  HEY365_WORKIQ_COMMAND    Work IQ CLI の起動コマンドを上書きする
  HEY365_WORKIQ_ACCOUNT    使用するアカウント（複数アカウント時）
  HEY365_LOG_LEVEL         silent|error|warn|info|debug
  HEY365_STATE_FILE        下書きの保存先
`;

interface Args {
  command: string;
  positionals: string[];
  flags: Record<string, string | boolean>;
}

function parseArgs(argv: string[]): Args {
  const flags: Record<string, string | boolean> = {};
  const positionals: string[] = [];

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index]!;
    if (token.startsWith('--')) {
      const key = token.slice(2);
      const next = argv[index + 1];
      if (next && !next.startsWith('--')) {
        flags[key] = next;
        index += 1;
      } else {
        flags[key] = true;
      }
    } else {
      positionals.push(token);
    }
  }

  const command = positionals.shift() ?? 'mcp';
  return { command, positionals, flags };
}

function numberFlag(flags: Args['flags'], key: string, fallback: number): number {
  const value = flags[key];
  const parsed = typeof value === 'string' ? Number(value) : NaN;
  return Number.isFinite(parsed) ? parsed : fallback;
}

function boolFlag(flags: Args['flags'], key: string): boolean {
  return flags[key] === true || flags[key] === 'true';
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));

  if (args.flags.help || args.flags.h || args.command === 'help') {
    process.stdout.write(USAGE);
    return;
  }
  if (args.flags.version || args.command === 'version') {
    process.stdout.write(`${HEY365_VERSION}\n`);
    return;
  }

  switch (args.command) {
    case 'mcp':
      await startStdioServer();
      return;

    case 'triage': {
      const response = await triageTool({
        hours: numberFlag(args.flags, 'hours', 36),
        limit: numberFlag(args.flags, 'limit', 10),
        includeDrafts: !boolFlag(args.flags, 'no-drafts'),
      });
      const text = response.content.map((part) => part.text).join('\n');
      if (response.isError) {
        process.stderr.write(`${text}\n`);
        process.exitCode = 1;
        return;
      }
      process.stdout.write(`${text}\n`);
      return;
    }

    case 'session': {
      const sessionId = args.positionals[0];
      if (!sessionId) {
        process.stderr.write('session id を指定してください。\n');
        process.exitCode = 1;
        return;
      }
      const summary = await summarizeSession(sessionId, { withDraft: args.flags['no-drafts'] !== true });
      process.stdout.write(`${formatSessionSummary(summary, timezone())}\n`);
      return;
    }

    case 'health': {
      const report = await checkHealth({ deep: args.flags.deep === true });
      process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
      process.exitCode = report.authenticated ? 0 : 1;
      return;
    }

    case 'setup': {
      const result = await runSetup({
        interactive: args.flags['no-interactive'] !== true,
        force: args.flags.force === true,
      });
      for (const step of result.steps) {
        process.stdout.write(`${step.ok ? '✅' : '❌'} ${step.step}${step.detail ? ` — ${step.detail}` : ''}\n`);
      }
      process.stdout.write(`\n${result.message}\n`);
      process.exitCode = result.ok ? 0 : 1;
      return;
    }

    case 'install': {
      const spec = buildServerSpec();
      const wantAll = args.flags.all === true;
      const ids = wantAll ? CLIENTS.map((client) => client.id) : args.positionals;

      if (ids.length === 0) {
        process.stdout.write('登録先クライアントを指定してください:\n\n');
        for (const client of CLIENTS) {
          process.stdout.write(`  ${client.id.padEnd(15)} ${client.label}\n`);
        }
        process.stdout.write('\n設定内容（mcpServers 形式）:\n\n');
        process.stdout.write(renderConfig('mcpServers', spec));
        return;
      }

      for (const id of ids) {
        const client = findClient(id);
        if (!client) {
          process.stderr.write(`未知のクライアント: ${id}\n`);
          process.exitCode = 1;
          continue;
        }
        const outcome = installForClient(client, spec, {
          dryRun: args.flags['dry-run'] === true,
          ...(typeof args.flags.path === 'string' ? { pathOverride: args.flags.path } : {}),
        });
        if (outcome.error) {
          process.stderr.write(`❌ ${outcome.label}: ${outcome.error}\n`);
          process.exitCode = 1;
        } else if (outcome.written) {
          process.stdout.write(`✅ ${outcome.label} → ${outcome.path}\n`);
        } else {
          process.stdout.write(`ℹ️ ${outcome.label} (${outcome.path})\n\n${outcome.preview}\n`);
        }
        if (client.note) process.stdout.write(`   ${client.note}\n`);
      }
      return;
    }

    default:
      process.stderr.write(`未知のコマンド: ${args.command}\n\n${USAGE}`);
      process.exitCode = 1;
  }
}

main().catch((error) => {
  const hey = asHey365Error(error);
  process.stderr.write(`${hey.nextStep('ja')}\n(code: ${hey.code})\n`);
  process.exitCode = 1;
});
