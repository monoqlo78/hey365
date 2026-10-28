/**
 * Hey365 error taxonomy (spec section 19).
 * Every user-facing failure maps to one of these codes so the MCP client can
 * present a concrete next action instead of a stack trace.
 */
export type Hey365ErrorCode =
  | 'WORKIQ_NOT_INSTALLED'
  | 'WORKIQ_NOT_AUTHENTICATED'
  | 'WORKIQ_AUTH_EXPIRED'
  | 'WORKIQ_ADMIN_CONSENT_REQUIRED'
  | 'WORKIQ_CONNECTION_ERROR'
  | 'WORKIQ_PERMISSION_DENIED'
  | 'WORKIQ_WRITE_DISABLED'
  | 'WORKIQ_EULA_REQUIRED'
  | 'WORKIQ_CREDENTIAL_STORE_ERROR'
  | 'SESSION_NOT_FOUND'
  | 'SEND_FAILED'
  | 'DRAFT_NOT_FOUND'
  | 'DRAFT_MISMATCH'
  | 'INVALID_INPUT'
  | 'SCHEDULE_FAILED'
  | 'INTERNAL_ERROR';

const NEXT_STEPS: Record<Hey365ErrorCode, { ja: string; en: string }> = {
  WORKIQ_NOT_INSTALLED: {
    ja: 'Work IQ CLI が見つかりません。`hey365_setup` を実行するか `npm i -g @microsoft/workiq` でインストールしてください。',
    en: 'Work IQ CLI was not found. Run `hey365_setup`, or install it with `npm i -g @microsoft/workiq`.',
  },
  WORKIQ_NOT_AUTHENTICATED: {
    ja: 'WorkIQ のサインインが必要です。`hey365_reconnect` が成功した場合だけ元の操作を再実行してください。ユーザー操作が必要な場合は、ターミナルで `hey365 setup` を実行してください。',
    en: 'WorkIQ needs sign-in. Retry the original operation only after `hey365_reconnect` succeeds. If user interaction is needed, run `hey365 setup` in a terminal.',
  },
  WORKIQ_AUTH_EXPIRED: {
    ja: 'WorkIQ の認証が期限切れです。`hey365_reconnect` が成功した場合だけ元の操作を再実行してください。再サインインが必要な場合はターミナルで `hey365 setup` を実行してください。',
    en: 'WorkIQ sign-in expired. Retry only after `hey365_reconnect` succeeds. If a new sign-in is required, run `hey365 setup` in a terminal.',
  },
  WORKIQ_ADMIN_CONSENT_REQUIRED: {
    ja: 'テナント管理者の同意が必要です。公式ガイド https://github.com/microsoft/work-iq/blob/main/ADMIN-INSTRUCTIONS.md に従い、管理者に WorkIQ のサービスプリンシパルと同意を確認してもらってください。',
    en: 'Tenant admin consent is required. Ask an admin to verify WorkIQ service principals and consent using https://github.com/microsoft/work-iq/blob/main/ADMIN-INSTRUCTIONS.md.',
  },
  WORKIQ_CONNECTION_ERROR: {
    ja: 'Microsoft 365 に接続できませんでした。ネットワークを確認して再実行してください。',
    en: 'Could not reach Microsoft 365. Check your network connection and retry.',
  },
  WORKIQ_PERMISSION_DENIED: {
    ja: 'WorkIQ の権限またはポリシーにより拒否されました。再ログインは繰り返さず、エラー詳細にあるスコープ・管理者同意・アクセス制御を確認してください。',
    en: 'WorkIQ denied this operation. Do not repeatedly sign in; check the reported scope, admin consent and access policy.',
  },
  WORKIQ_CREDENTIAL_STORE_ERROR: {
    ja: 'OS の資格情報ストアまたは暗号化キャッシュを利用できません。資格情報ストアへのアクセスを確認してください。平文保存には切り替えません。',
    en: 'The OS credential store or encrypted token cache is unavailable. Check credential-store access; plaintext storage is not used.',
  },
  WORKIQ_WRITE_DISABLED: {
    ja: '送信権限がありません。要約と返信案の作成までは利用できます。',
    en: 'Sending is not permitted for this account. Triage and draft generation still work.',
  },
  WORKIQ_EULA_REQUIRED: {
    ja: 'Work IQ の使用許諾への同意が必要です。`npx @microsoft/workiq accept-eula` を実行してください。',
    en: 'You must accept the Work IQ EULA. Run `npx @microsoft/workiq accept-eula`.',
  },
  SESSION_NOT_FOUND: {
    ja: '指定された session / 会議 / 会話が見つかりませんでした。件名や日付で検索し直してください。',
    en: 'The requested session, meeting, or conversation was not found. Try searching by subject or date.',
  },
  SEND_FAILED: {
    ja: '送信に失敗しました。対象の会話がまだ存在するか確認してください。',
    en: 'Sending failed. Verify that the target conversation still exists.',
  },
  DRAFT_NOT_FOUND: {
    ja: '指定された番号の返信案がありません。先に `hey365` を実行してください。',
    en: 'No draft exists for that number. Run `hey365` first.',
  },
  DRAFT_MISMATCH: {
    ja: '返信案が更新されています。内容を確認してから送信してください。',
    en: 'The draft changed since it was shown. Review it again before sending.',
  },
  INVALID_INPUT: {
    ja: '指定された値を解釈できませんでした。日付は `YYYY-MM-DD` 形式で指定してください。',
    en: 'That value could not be interpreted. Dates must be given as `YYYY-MM-DD`.',
  },
  SCHEDULE_FAILED: {
    ja: '定期実行の登録に失敗しました。管理者権限で実行しているか、表示されたコマンドを手動で実行できるか確認してください。',
    en: 'Could not register the scheduled run. Check your permissions, or run the printed command manually.',
  },
  INTERNAL_ERROR: {
    ja: 'Hey365 の内部エラーです。`hey365_health` で状態を確認してください。',
    en: 'Internal Hey365 error. Run `hey365_health` to inspect the current state.',
  },
};

export class Hey365Error extends Error {
  readonly code: Hey365ErrorCode;
  readonly detail?: string;

  constructor(code: Hey365ErrorCode, message?: string, detail?: string) {
    super(message ?? NEXT_STEPS[code].en);
    this.name = 'Hey365Error';
    this.code = code;
    this.detail = detail;
  }

  nextStep(locale: 'ja' | 'en' = 'ja'): string {
    return NEXT_STEPS[this.code][locale];
  }

  toPayload(): {
    code: Hey365ErrorCode;
    message: string;
    nextStep: string;
    nextStepEn: string;
    detail?: string;
    recovery?: RecoveryHint;
  } {
    return {
      code: this.code,
      message: this.message,
      nextStep: this.nextStep('ja'),
      nextStepEn: this.nextStep('en'),
      ...(this.detail ? { detail: this.detail } : {}),
      ...(recoveryFor(this.code) ? { recovery: recoveryFor(this.code) as RecoveryHint } : {}),
    };
  }
}

/**
 * Machine-readable recovery instruction. Clients that only read
 * `structuredContent` still learn that the failure is self-healing and which
 * tool fixes it, so they stop bouncing the problem back to the user.
 */
export interface RecoveryHint {
  /** Tool the client should call before retrying. */
  tool: 'hey365_reconnect';
  /** Hey365 already attempted this automatically before failing. */
  automatic: boolean;
  /** Re-running the original tool unchanged is expected to succeed. */
  retryOriginalCall: boolean;
  /** Recovery needs no user confirmation. */
  requiresUserConfirmation: false;
}

/** Auth failures Hey365 can repair by itself. */
export function isAuthFailure(code: Hey365ErrorCode): boolean {
  return code === 'WORKIQ_NOT_AUTHENTICATED' || code === 'WORKIQ_AUTH_EXPIRED';
}

export function recoveryFor(code: Hey365ErrorCode): RecoveryHint | undefined {
  if (!isAuthFailure(code)) return undefined;
  return {
    tool: 'hey365_reconnect',
    automatic: true,
    retryOriginalCall: true,
    requiresUserConfirmation: false,
  };
}

export function asHey365Error(error: unknown): Hey365Error {
  if (error instanceof Hey365Error) return error;
  const message = error instanceof Error ? error.message : String(error);
  return new Hey365Error('INTERNAL_ERROR', message);
}

/**
 * Maps raw Work IQ / Microsoft Graph failure text onto the Hey365 taxonomy.
 * Matching is intentionally broad because the CLI localises its own messages.
 */
export function classifyFailure(raw: string, statusCode?: number): Hey365ErrorCode {
  const text = raw.toLowerCase();

  if (
    text.includes('enoent') ||
    text.includes('command not found') ||
    text.includes('is not recognized') ||
    text.includes('could not determine executable') ||
    text.includes('npm error 404')
  ) {
    return 'WORKIQ_NOT_INSTALLED';
  }

  if (text.includes('eula')) return 'WORKIQ_EULA_REQUIRED';

  if (
    text.includes('aadsts65001') ||
    text.includes('aadsts650052') ||
    text.includes('consent_required') ||
    text.includes('admin consent') ||
    text.includes('auth consent')
  ) {
    return 'WORKIQ_ADMIN_CONSENT_REQUIRED';
  }

  if (
    statusCode === 403 ||
    /\b(?:http(?: status)?\s*[:=]?\s*403|statuscode["'\s:]+403)\b/.test(text) ||
    text.includes('accessdenied') ||
    text.includes('access denied') ||
    text.includes('authorization_requestdenied') ||
    text.includes('missing scope permissions') ||
    text.includes('insufficient privileges') ||
    text.includes('forbidden')
  ) {
    return 'WORKIQ_PERMISSION_DENIED';
  }

  if (
    text.includes('aadsts50076') ||
    text.includes('aadsts50079') ||
    text.includes('aadsts700082') ||
    text.includes('aadsts700081') ||
    text.includes('aadsts50173') ||
    text.includes('interaction_required') ||
    text.includes('msaluirequired') ||
    text.includes('invalid_grant') ||
    text.includes('token expired') ||
    text.includes('token is expired') ||
    text.includes('refresh token has expired')
  ) {
    return 'WORKIQ_AUTH_EXPIRED';
  }

  if (
    text.includes('no cached account') ||
    text.includes('no account found') ||
    text.includes('not signed in') ||
    text.includes('please log in') ||
    text.includes('please sign in') ||
    text.includes('auth login') ||
    text.includes('unauthenticated') ||
    text.includes('invalidauthenticationtoken') ||
    text.includes('failed to acquire token') ||
    statusCode === 401
  ) {
    return 'WORKIQ_NOT_AUTHENTICATED';
  }

  if (
    text.includes('etimedout') ||
    text.includes('econnreset') ||
    text.includes('enotfound') ||
    text.includes('socket hang up') ||
    text.includes('fetch failed') ||
    text.includes('timed out') ||
    text.includes('econnrefused') ||
    text.includes('network') ||
    (statusCode !== undefined && statusCode >= 500)
  ) {
    return 'WORKIQ_CONNECTION_ERROR';
  }

  return 'INTERNAL_ERROR';
}
