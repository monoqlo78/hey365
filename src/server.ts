import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

import {
  digestInputSchema,
  digestTool,
  draftInputSchema,
  draftTool,
  findInputSchema,
  findTool,
  healthInputSchema,
  healthTool,
  installInputSchema,
  installTool,
  mutesInputSchema,
  mutesTool,
  sendInputSchema,
  sendTool,
  sessionInputSchema,
  sessionTool,
  setupInputSchema,
  setupTool,
  snoozeInputSchema,
  snoozeTool,
  scheduleInputSchema,
  scheduleTool,
  triageInputSchema,
  triageTool,
} from './tools/index.js';
import { logger } from './utils/logger.js';

export const HEY365_VERSION = '0.1.0';

const INSTRUCTIONS = `Hey365 は Microsoft 365 (Outlook / Teams / Calendar) を横断して
「あなたがまだ返信しておらず、あなたが対応すべき用件」だけを抽出し、返信案を作成します。

ルーティング:
- 「Hey365」          -> hey365_triage (hours=36)
- 「Hey365 24時間」    -> hey365_triage (hours=24)
- 「Hey365 3営業日」「先週の金曜を基準に」 -> hey365_triage (businessDays / asOf)
- 「今日のまとめ」「朝のダイジェスト」 -> hey365_digest
- 「Hey365 session X」 -> hey365_session (sessionId=X)
- 「このsessionをまとめて X」 -> hey365_session
- 「1をもう少し柔らかく」「1を英語にして」「1を短く」 -> hey365_draft
- 「1を送って」「1と3を送って」「全部送って」 -> hey365_send
- 「1は明日でいい」「1を後回し」 -> hey365_snooze (action=snooze)
- 「1は対応済み」「1はもう終わってる」 -> hey365_snooze (action=done)
- 「1をまた表示して」 -> hey365_snooze (action=unmute)
- 「何を隠してたっけ」 -> hey365_mutes
- 「毎朝8時にダイジェストを出して」 -> hey365_schedule (action=add)
- 「Xの件どうなってる?」「Yさんとのやりとりを見せて」 -> hey365_find

重要: hey365_triage は下書きまでしか作りません。ユーザーが明示的に送信を指示した場合のみ
hey365_send を呼び出してください。ユーザーの確認なしに送信してはいけません。`;

export function createServer(): McpServer {
  const server = new McpServer(
    { name: 'hey365', version: HEY365_VERSION },
    { instructions: INSTRUCTIONS, capabilities: { tools: {}, logging: {} } },
  );

  // Primary entry point. Registered twice so that both "hey365" and
  // "hey365_triage" resolve, which is what the natural-language routing in
  // section 14 of the spec expects.
  server.registerTool(
    'hey365',
    {
      title: 'Hey365 triage',
      description:
        '過去N時間(既定36)の Outlook / Teams / 会議を調査し、自分が返信・対応すべき用件だけを抽出して返信案を作成する。送信はしない。',
      inputSchema: triageInputSchema,
    },
    triageTool,
  );

  server.registerTool(
    'hey365_triage',
    {
      title: 'Hey365 triage (alias)',
      description: 'hey365 と同じ。過去N時間の M365 を調査して対応が必要な項目を抽出する。',
      inputSchema: triageInputSchema,
    },
    triageTool,
  );

  server.registerTool(
    'hey365_session',
    {
      title: 'Hey365 session summary',
      description:
        '会議 / Teams スレッド / Outlook 会話を要約し、決定事項・Action Items・自分の宿題・未解決事項・返信要否を出力する。',
      inputSchema: sessionInputSchema,
    },
    sessionTool,
  );

  server.registerTool(
    'hey365_draft',
    {
      title: 'Hey365 refine draft',
      description: '既に提示した返信案を自然言語の指示で修正する（柔らかく / 英語に / 短く など）。',
      inputSchema: draftInputSchema,
    },
    draftTool,
  );

  server.registerTool(
    'hey365_send',
    {
      title: 'Hey365 send replies',
      description:
        'ユーザーが明示的に指示した返信案だけを、元のスレッドへの返信として送信する。新規スレッドは作らない。',
      inputSchema: sendInputSchema,
    },
    sendTool,
  );

  server.registerTool(
    'hey365_digest',
    {
      title: 'Hey365 daily digest',
      description:
        '今日の会議・期限が来ている用件・放置中の用件・未返信を1画面にまとめる。朝いちばんの確認用。送信はしない。',
      inputSchema: digestInputSchema,
    },
    digestTool,
  );

  server.registerTool(
    'hey365_snooze',
    {
      title: 'Hey365 snooze / done',
      description:
        '提示した項目を一定期間隠す(snooze)、対応済みにする(done)、再表示する(unmute)。新しいメッセージが届けば自動で再表示される。',
      inputSchema: snoozeInputSchema,
    },
    snoozeTool,
  );

  server.registerTool(
    'hey365_mutes',
    {
      title: 'Hey365 list hidden items',
      description: '現在 snooze / 対応済みで隠している項目の一覧を表示する。',
      inputSchema: mutesInputSchema,
    },
    mutesTool,
  );

  server.registerTool(
    'hey365_find',
    {
      title: 'Hey365 find conversations',
      description:
        '人名・案件名・製品名などで Outlook と Teams を横断検索し、それぞれの会話で誰が最後に発言したか（自分待ちか相手待ちか）を表示する。時間の窓に関係なく探せる。',
      inputSchema: findInputSchema,
    },
    findTool,
  );

  server.registerTool(
    'hey365_schedule',
    {
      title: 'Hey365 scheduled runs',
      description:
        'ダイジェストの定期実行を OS のスケジューラ(schtasks / crontab)に登録・解除・一覧する。既定では内容を表示するだけで、apply=true を指定したときだけ実際に登録する。',
      inputSchema: scheduleInputSchema,
    },
    scheduleTool,
  );

  server.registerTool(
    'hey365_health',
    {
      title: 'Hey365 health check',
      description: 'Work IQ 接続・認証・読み取り/書き込み権限・タイムゾーンを確認する。',
      inputSchema: healthInputSchema,
    },
    healthTool,
  );

  server.registerTool(
    'hey365_setup',
    {
      title: 'Hey365 setup / auth recovery',
      description:
        'Work IQ が未インストール・未認証・WAM 認証失敗の場合に、インストールと Broker 無効化＋ブラウザ認証を実行して復旧する。',
      inputSchema: setupInputSchema,
    },
    setupTool,
  );

  server.registerTool(
    'hey365_install',
    {
      title: 'Hey365 client install',
      description:
        'GitHub Copilot / Claude Code / Codex / Cursor / Windsurf / Scout などの MCP 設定に Hey365 を登録する。',
      inputSchema: installInputSchema,
    },
    installTool,
  );

  return server;
}

export async function startStdioServer(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  logger.info(`hey365 MCP server ${HEY365_VERSION} ready (stdio)`);
}
