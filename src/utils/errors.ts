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
    ja: 'Work IQ の認証が必要です。`hey365_setup` を実行するとブラウザ認証を開始します。',
    en: 'Work IQ needs to sign in. Run `hey365_setup` to start browser-based authentication.',
  },
  WORKIQ_AUTH_EXPIRED: {
    ja: 'Work IQ の認証が期限切れです。`hey365_setup` で再ログインしてください。',
    en: 'Your Work IQ session expired. Run `hey365_setup` to sign in again.',
  },
  WORKIQ_ADMIN_CONSENT_REQUIRED: {
    ja: 'テナント管理者の同意が必要です。管理者に `npx @microsoft/workiq auth consent` の実行を依頼してください。',
    en: 'Tenant admin consent is required. Ask an admin to run `npx @microsoft/workiq auth consent`.',
  },
  WORKIQ_CONNECTION_ERROR: {
    ja: 'Microsoft 365 に接続できませんでした。ネットワークを確認して再実行してください。',
    en: 'Could not reach Microsoft 365. Check your network connection and retry.',
  },
  WORKIQ_PERMISSION_DENIED: {
    ja: 'このデータを読み取る権限がありません。取得できた範囲で続行します。',
    en: 'You do not have permission to read this data. Continuing with what is available.',
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
  } {
    return {
      code: this.code,
      message: this.message,
      nextStep: this.nextStep('ja'),
      nextStepEn: this.nextStep('en'),
      ...(this.detail ? { detail: this.detail } : {}),
    };
  }
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
    text.includes('aadsts50076') ||
    text.includes('aadsts50079') ||
    text.includes('aadsts700082') ||
    text.includes('interaction_required') ||
    text.includes('invalid_grant') ||
    text.includes('token expired') ||
    text.includes('token is expired')
  ) {
    return 'WORKIQ_AUTH_EXPIRED';
  }

  if (
    text.includes('aadsts65001') ||
    text.includes('consent_required') ||
    text.includes('admin consent') ||
    text.includes('auth consent')
  ) {
    return 'WORKIQ_ADMIN_CONSENT_REQUIRED';
  }

  if (
    text.includes('no cached account') ||
    text.includes('not signed in') ||
    text.includes('please log in') ||
    text.includes('auth login') ||
    text.includes('unauthenticated') ||
    statusCode === 401
  ) {
    return 'WORKIQ_NOT_AUTHENTICATED';
  }

  if (statusCode === 403 || text.includes('accessdenied') || text.includes('forbidden')) {
    return 'WORKIQ_PERMISSION_DENIED';
  }

  if (
    text.includes('etimedout') ||
    text.includes('econnreset') ||
    text.includes('enotfound') ||
    text.includes('socket hang up') ||
    text.includes('network') ||
    (statusCode !== undefined && statusCode >= 500)
  ) {
    return 'WORKIQ_CONNECTION_ERROR';
  }

  return 'INTERNAL_ERROR';
}
